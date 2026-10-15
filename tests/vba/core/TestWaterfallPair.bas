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

' Deco_SetLegend (cascata comparativa) grava a opcao "legend" antes de aplicar; se aplicar falhar, a opcao volta ao valor anterior e o erro chega
' ao chamador (contrato de ApplyOption: aplica e SO ENTAO grava). Gancho wfp-legend:right dentro de Wfp_SetLegend; zerado em todos os caminhos.
Public Sub Test_Wfp_SetLegend_FailRestoresPreviousOption()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, en As Long, es As String, ed As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    On Error GoTo Fail
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    cht.Refresh
    AssertEqual "inbar", Registry_GetOption(cht, "legend"), "o padrao e 'dentro das barras'"
    Set s = Chart_FindSeries(cht, "tqq_names")
    AssertEqual "Segmento 1", s.Points(1).DataLabel.Text, "os nomes estao nas barras antes"
    Util_FailThinSyncForTest = "wfp-legend:right"
    On Error Resume Next
    Deco_SetLegend cht, "right"
    en = Err.Number
    On Error GoTo Fail
    Util_FailThinSyncForTest = ""
    AssertEqual vbObjectError + 9996, en, "o erro chega ao chamador"
    AssertEqual "inbar", Registry_GetOption(cht, "legend"), "a opcao continua a anterior"
    AssertEqual 1, Util_ToDouble(Registry_Block(cht).Cells(1, WFP_META_NAMEON).Value), "os nomes continuam gerenciados"
    cht.Refresh
    AssertEqual "Segmento 1", s.Points(1).DataLabel.Text, "os nomes continuam nas barras"
    AssertFalse cht.HasLegend, "sem legenda, como antes"
    Deco_SetLegend cht, "right"                      ' sem o gancho a mesma troca funciona
    cht.Refresh
    AssertEqual "right", Registry_GetOption(cht, "legend")
    AssertTrue cht.HasLegend: AssertEqual "", s.Points(1).DataLabel.Text, "Direita: sem nomes nas barras"
    Util_FailThinSyncForTest = ""
    Teardown wb
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailThinSyncForTest = ""
    Teardown wb
    Err.Raise en, es, ed
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

' Estilo dos conectores: tracejado cinza 115 fino, sem ponta e sem marcador (meio-termo entre a barra escura da serie 1 e a clara da serie 2).
Private Sub AssertConnectorStyle(s As Series, label As String)
    AssertTrue s.HasErrorBars, label & ": barra de erro"
    AssertEqual xlNoCap, s.ErrorBars.EndStyle, label & ": sem ponta"
    AssertEqual msoTrue, s.ErrorBars.Format.Line.Visible, label & ": visivel"
    AssertEqual RGB(115, 115, 115), s.ErrorBars.Format.Line.ForeColor.RGB, label & ": cinza 115"
    AssertEqual msoLineDash, s.ErrorBars.Format.Line.DashStyle, label & ": tracejada"
    AssertAlmost 0.75, s.ErrorBars.Format.Line.Weight, 0.001, label & ": 0,75 pt"
    AssertEqual xlMarkerStyleNone, s.MarkerStyle, label & ": sem marcador"
    AssertEqual xlPrimary, s.AxisGroup, label & ": eixo principal"
End Sub

' Indice (em SeriesCollection) da serie de colunas que desenha a barra i do lado j: a parte positiva (2 e 4) ou, se ela tem altura zero, a negativa
' (3 e 5): a geometria de um ponto de altura zero nao e confiavel, a da parte que existe sim.
Private Function BarSeriesIdx(cht As Chart, i As Long, j As Long) As Long
    BarSeriesIdx = IIf(j = 1, 2, 4)
    If SideVal(cht, i, j, "PP") = 0 Then BarSeriesIdx = BarSeriesIdx + 1
End Function

' O Excel NAO expoe a geometria de uma barra de erro (nem Left/Width/Top do traco): o que se confere e o ponto XY de cada conector (Top = altura
' do fim da barra i, regra de tres sobre o eixo real) e a MATEMATICA da barra de erro a partir das celulas do bloco (x + CLen cai na borda
' esquerda da barra i + 1 da MESMA serie, regra de tres sobre a area interna), nao o desenho; o desenho e conferido nos PNG da verificacao e o
' XML salvo (Test_Wfp_Connectors_SavedXml_HoldsErrBarsOnBlock) garante que a barra de erro cita a faixa do bloco.
Private Sub AssertConnectorsGeometry(cht As Chart)
    Dim body As Range, ca As Series, cb As Series, cs As Series, i As Long, j As Long, n As Long, x As Double, clen As Double, hw As Double
    Dim p As Point, q As Point
    Set ca = Chart_FindSeries(cht, "tqq_conn_a"): Set cb = Chart_FindSeries(cht, "tqq_conn_b")
    AssertTrue Not ca Is Nothing, "conector da serie 1": AssertTrue Not cb Is Nothing, "conector da serie 2"
    n = Wfp_StepCount(cht)
    AssertEqual n - 1, ca.Points.Count, "n - 1 conectores": AssertEqual n - 1, cb.Points.Count
    Set body = BodyOf(cht)
    hw = Util_ToDouble(Registry_Block(cht).Cells(1, WFP_META_HW).Value)
    ForceLayout cht
    For i = 1 To n - 1
        clen = Util_ToDouble(body.Cells(i + 1, WFP_COL_CLEN).Value)
        AssertAlmost 3 - 2 * hw, clen, 0.000001, "comprimento do conector " & i
        For j = 1 To 2
            If j = 1 Then
                Set cs = ca: x = Util_ToDouble(body.Cells(i + 1, WFP_COL_CXA).Value)
            Else
                Set cs = cb: x = Util_ToDouble(body.Cells(i + 1, WFP_COL_CXB).Value)
            End If
            AssertAlmost Wfp_SlotOf(i, j) + hw, x, 0.000001, "x do conector " & i & "/" & j
            Set p = cht.SeriesCollection(BarSeriesIdx(cht, i, j)).Points(Wfp_SlotOf(i, j))
            Set q = cht.SeriesCollection(BarSeriesIdx(cht, i + 1, j)).Points(Wfp_SlotOf(i + 1, j))
            AssertAlmost p.Left + p.Width, XOfX(cht, x), 1, "sai da borda direita da barra " & i & "/" & j
            AssertAlmost q.Left, XOfX(cht, x + clen), 1, "chega a borda esquerda da proxima barra da mesma serie " & i & "/" & j
            AssertAlmost YOf(cht, SideVal(cht, i, j, "End")), cs.Points(i).Top, 1, "na altura do fim da barra " & i & "/" & j
        Next
    Next
End Sub

' Review Focus 5 (conectores): do lado direito da barra i ao lado esquerdo da barra i + 1 da MESMA serie, na altura do fim da barra i; vivos.
Public Sub Test_Wfp_Connectors_GeometryAndLive()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    AssertEqual "1", Registry_GetOption(cht, "connectors")
    AssertEqual 4, Chart_FindSeries(cht, "tqq_conn_a").Points.Count, "n - 1 conectores"
    AssertConnectorStyle Chart_FindSeries(cht, "tqq_conn_a"), "serie 1"
    AssertConnectorStyle Chart_FindSeries(cht, "tqq_conn_b"), "serie 2"
    AssertConnectorsGeometry cht
    Set body = BodyOf(cht)
    ws.Range("B3").Value = -9
    cht.Refresh
    AssertEqual 18, Util_ToDouble(body.Cells(3, WFP_COL_CYA).Value), "sem o ThinQQ o conector acompanha o dado"
    CloseTemp wb
End Sub

' Cruza o zero: fim negativo (abaixo do zero do eixo), barras so de parte negativa (a positiva de altura zero) e total negativo.
Public Sub Test_Wfp_Connectors_GeometryCrossesZero()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = FixtureZero(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    AssertTrue SideVal(cht, 2, 1, "End") < 0, "a serie A termina a etapa 2 abaixo de zero"
    AssertTrue SideVal(cht, 3, 1, "PP") = 0, "a barra 3 da serie A so tem parte negativa (positiva de altura zero)"
    AssertConnectorStyle Chart_FindSeries(cht, "tqq_conn_a"), "serie 1"
    AssertConnectorsGeometry cht
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Connectors_ToggleGapLegendAndSingleStep()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, c1 As Chart, body As Range, x As Double, clen As Double, p As Point, q As Point
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    d.Source.Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart()
    Core_SetToggle "connectors", False
    AssertTrue Chart_FindSeries(cht, "tqq_conn_a") Is Nothing: AssertTrue Chart_FindSeries(cht, "tqq_conn_b") Is Nothing
    AssertEqual "0", Registry_GetOption(cht, "connectors"): AssertFalse Core_GetToggle("connectors")
    Core_SetToggle "connectors", True
    AssertEqual 4, Chart_FindSeries(cht, "tqq_conn_a").Points.Count
    Core_SetOption "gap", "80"
    ForceLayout cht
    Set body = BodyOf(cht)
    x = Util_ToDouble(body.Cells(2, WFP_COL_CXA).Value): clen = Util_ToDouble(body.Cells(2, WFP_COL_CLEN).Value)
    Set p = cht.SeriesCollection(2).Points(1): Set q = cht.SeriesCollection(2).Points(4)
    AssertAlmost p.Left + p.Width, XOfX(cht, x), 1, "espaco 80: o conector continua na borda da barra"
    AssertAlmost q.Left, XOfX(cht, x + clen), 1
    x = Util_ToDouble(body.Cells(2, WFP_COL_CXB).Value)
    Set p = cht.SeriesCollection(4).Points(Wfp_SlotOf(1, 2)): Set q = cht.SeriesCollection(4).Points(Wfp_SlotOf(2, 2))
    AssertAlmost p.Left + p.Width, XOfX(cht, x), 1, "espaco 80, serie 2: o conector continua na borda da barra"
    AssertAlmost q.Left, XOfX(cht, x + clen), 1, "espaco 80, serie 2: chega a borda da proxima barra"
    AssertConnectorsGeometry cht
    Core_SetOption "legend", "right"
    AssertEqual 2, cht.Legend.LegendEntries.Count, "os conectores ficam fora da legenda"
    Core_SetToggle "connectors", False
    AssertEqual 2, cht.Legend.LegendEntries.Count, "desligar com a legenda visivel: continua com 2 entradas"
    Core_SetToggle "connectors", True
    AssertEqual 4, Chart_FindSeries(cht, "tqq_conn_b").Points.Count
    AssertEqual 2, cht.Legend.LegendEntries.Count, "ligar com a legenda visivel: os conectores continuam fora da legenda"
    ws.Range("E1:G1").Value = Array("Etapa", "S1", "S2"): ws.Range("E2:G2").Value = Array("Unico", 5, 7)
    Set c1 = Wfp_Insert(ws, ChartData_FromRange(ws.Range("E1:G2"))).Chart
    AssertEqual 2, SlotN(c1), "uma etapa: 2 fatias, sem vao"
    AssertTrue Chart_FindSeries(c1, "tqq_conn_a") Is Nothing, "uma etapa: sem conectores"
    AssertEqual "1", Registry_GetOption(c1, "connectors"), "a opcao fica gravada"
    Teardown wb
End Sub

' Falha no meio de ligar os conectores (gancho depois da barra de erro da 2a serie, depois da 1a serie e depois do passo da legenda): nada fica
' pela metade (nenhuma das duas series, opcao "0" mesmo que antes fosse "1", mesmo numero de series do grafico, legenda com 2 entradas) e o erro
' chega ao chamador. Sem o gancho, ligar de novo funciona. O gancho e zerado no mesmo caminho, antes de qualquer afirmacao, e no tratamento de erro.
Public Sub Test_Wfp_Connectors_FailMidway_RollsBack()
    Dim wb As Workbook, d As clsChartData, cht As Chart, en As Long, es As String, ed As String, base As Long, hook As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    On Error GoTo Fail
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "1", Registry_GetOption(cht, "connectors"), "ligados desde o Insert: DropConnectors e o desfazer rodam"
    base = cht.SeriesCollection.Count
    For Each hook In Array("wfp-conn:2", "wfp-conn:1")
        Util_FailThinSyncForTest = CStr(hook)
        On Error Resume Next
        Wfp_SetConnectors cht, True
        en = Err.Number
        On Error GoTo Fail
        Util_FailThinSyncForTest = ""
        AssertEqual vbObjectError + 9996, en, hook & ": o erro chega ao chamador"
        AssertTrue Chart_FindSeries(cht, "tqq_conn_a") Is Nothing, hook & ": sem o conector da serie 1"
        AssertTrue Chart_FindSeries(cht, "tqq_conn_b") Is Nothing, hook & ": sem o conector da serie 2"
        AssertEqual "0", Registry_GetOption(cht, "connectors"), hook & ": a opcao reflete o estado real"
        AssertEqual base - 2, cht.SeriesCollection.Count, hook & ": so as duas series dos conectores sairam"
        Wfp_SetConnectors cht, True
        AssertEqual "1", Registry_GetOption(cht, "connectors"), hook & ": ligar de novo funciona"
        AssertEqual base, cht.SeriesCollection.Count
        AssertConnectorStyle Chart_FindSeries(cht, "tqq_conn_a"), "religado a"
        AssertConnectorStyle Chart_FindSeries(cht, "tqq_conn_b"), "religado b"
    Next
    ' conectores desligados antes: a falha deixa tudo como estava (opcao "0", nenhuma serie)
    Wfp_SetConnectors cht, False
    base = cht.SeriesCollection.Count
    Util_FailThinSyncForTest = "wfp-conn:2"
    On Error Resume Next
    Wfp_SetConnectors cht, True
    en = Err.Number
    On Error GoTo Fail
    Util_FailThinSyncForTest = ""
    AssertEqual vbObjectError + 9996, en, "desligados antes: o erro chega ao chamador"
    AssertEqual "0", Registry_GetOption(cht, "connectors")
    AssertEqual base, cht.SeriesCollection.Count, "desligados antes: nenhuma serie sobra"
    ' com a legenda visivel: falha DEPOIS das duas series, no passo da legenda: o desfazer cobre as duas
    Wfp_SetLegend cht, "right"
    AssertEqual 2, cht.Legend.LegendEntries.Count
    Util_FailThinSyncForTest = "wfp-connlegend:1"
    On Error Resume Next
    Wfp_SetConnectors cht, True
    en = Err.Number
    On Error GoTo Fail
    Util_FailThinSyncForTest = ""
    AssertEqual vbObjectError + 9996, en, "legenda: o erro chega ao chamador"
    AssertTrue Chart_FindSeries(cht, "tqq_conn_a") Is Nothing, "legenda: sem o conector da serie 1"
    AssertTrue Chart_FindSeries(cht, "tqq_conn_b") Is Nothing, "legenda: sem o conector da serie 2"
    AssertEqual "0", Registry_GetOption(cht, "connectors"), "legenda: opcao atual"
    AssertEqual base, cht.SeriesCollection.Count
    AssertEqual 2, cht.Legend.LegendEntries.Count, "legenda: 2 entradas"
    Wfp_SetConnectors cht, True
    AssertEqual "1", Registry_GetOption(cht, "connectors")
    AssertEqual 2, cht.Legend.LegendEntries.Count, "legenda: os conectores ficam fora"
    Util_FailThinSyncForTest = ""
    CloseTemp wb
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailThinSyncForTest = ""
    CloseTemp wb
    Err.Raise en, es, ed
End Sub

' O arquivo salvo guarda a barra de erro dos conectores citando a faixa do bloco (c:errBars/c:plus/c:numRef com a planilha _ThinQQ), nao uma
' constante (c:numLit): a prova de que continuam vivos sem o ThinQQ. Salva uma copia na pasta temporaria do teste e extrai o XML do grafico
' pela pasta compactada do Windows (Shell.Application, dentro do proprio Excel), sem tocar na pasta do usuario. NUNCA inicia processo filho
' (PowerShell/cmd): o antivirus/EDR da maquina le "Excel que salvou/vinculou uma pasta .xlsx e depois roda powershell.exe" como documento
' malicioso e ENCERRA o Excel da suite (0x800706BE na execucao completa do core; o Test-VbaLint barra a volta).
Public Sub Test_Wfp_Connectors_SavedXml_HoldsErrBarsOnBlock()
    Dim wb As Workbook, d As clsChartData, cht As Chart, dir As String, xlsx As String, zip As String, outDir As String, xml As String, pos As Long, endPos As Long
    Dim fso As Object, ts As Object, sh As Object, src As Object, it As Object, f As Object, barXml As String, k As Long, q As String
    Dim n As Long, want As Double, got As Double, t0 As Single
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    dir = TempDir(): xlsx = dir & "\copia.xlsx": zip = dir & "\copia.zip": outDir = dir & "\xml"
    wb.SaveCopyAs xlsx
    CloseTemp wb
    q = Chr$(34)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile xlsx, zip                               ' a pasta compactada do Windows so abre pela extensao .zip
    fso.CreateFolder outDir
    Set sh = CreateObject("Shell.Application")
    Set src = sh.Namespace(CVar(zip & "\xl\charts"))
    AssertTrue Not src Is Nothing, "pasta xl\charts do arquivo salvo"
    For Each it In src.Items
        If Not it.IsFolder Then
            If LCase$(fso.GetFileName(it.Path)) Like "chart*.xml" Then
                want = want + it.Size
                n = n + 1
                sh.Namespace(CVar(outDir)).CopyHere it, 4 + 16 + 512 + 1024     ' sem progresso, sim para tudo, sem perguntas nem erro na tela
            End If
        End If
    Next
    AssertTrue n > 0, "XML do grafico no arquivo salvo"
    ' CopyHere e assincrono: espera os n arquivos com o tamanho descompactado completo (teto de 15 s)
    t0 = Timer
    Do
        got = 0
        For Each f In fso.GetFolder(outDir).Files
            got = got + f.Size
        Next
        If fso.GetFolder(outDir).Files.Count >= n And got >= want Then Exit Do
        If Abs(Timer - t0) > 15 Then Exit Do
        DoEvents
    Loop
    AssertEqual n, fso.GetFolder(outDir).Files.Count, "XML do grafico extraido"
    For Each f In fso.GetFolder(outDir).Files
        Set ts = fso.OpenTextFile(f.Path, 1, False, 0)
        If Not ts.AtEndOfStream Then xml = xml & ts.ReadAll
        ts.Close
    Next
    AssertTrue Len(xml) > 1000, "XML do grafico lido"
    pos = 1
    Do
        pos = InStr(pos, xml, "<c:errBars>")
        If pos = 0 Then Exit Do
        k = k + 1
        endPos = InStr(pos, xml, "</c:errBars>")
        AssertTrue endPos > pos, "errBars fechada"
        barXml = Mid$(xml, pos, endPos - pos)
        AssertTrue InStr(barXml, "<c:errDir val=" & q & "x" & q) > 0, "barra de erro na direcao X (" & k & ")"
        AssertTrue InStr(barXml, "<c:plus><c:numRef>") > 0, "o plus e uma referencia (" & k & ")"
        AssertTrue InStr(barXml, "<c:numLit>") = 0, "nenhuma constante (" & k & ")"
        AssertTrue InStr(barXml, CALC_SHEET) > 0, "a referencia cita a planilha do bloco (" & k & ")"
        pos = endPos
    Loop
    AssertEqual 2, k, "uma barra de erro por serie de conectores"
End Sub

' Copia do grafico (como colada: outro nome) com o original apagado: a limpeza nao remove o bloco (colunas, XY e rotulos citam o bloco).
Public Sub Test_Wfp_CopiedChart_KeepsBlockAlive()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, dup As Object
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Wfp_Insert ws, d
    Set dup = ws.ChartObjects(1).Duplicate
    ws.ChartObjects(2).Name = "Gráfico 2"
    ws.ChartObjects(1).Delete
    AssertEqual 0, CalcSheet_Cleanup(wb), "a copia cita o bloco: nada e removido"
    AssertFalse CalcSheet_CleanupReadFailed()
    AssertEqual 1, TqqNameCount(wb)
    CloseTemp wb
End Sub

' ROUND do Excel (metade para longe do zero), nao o Round do VBA (banqueiro).
Private Function XlRound(x As Double) As Double
    XlRound = Sgn(x) * Int(Abs(x) + 0.5)
End Function

' Texto esperado da seta: (B - A) / A em % inteiro com sinal; "n/d" com A = 0 ou sinais opostos.
Private Function PctText(a As Double, b As Double) As String
    Dim r As Double
    If a = 0 Or a * b < 0 Then PctText = "n/d": Exit Function
    r = XlRound(100 * (b - a) / a)
    PctText = IIf(r > 0, "+", "") & Format$(r, "0") & "%"
End Function

Private Function ArrowLabel(cht As Chart, i As Long) As String
    ArrowLabel = Chart_FindSeries(cht, "tqq_pd").Points(WFP_ROWS_PER_PAIR * (i - 1) + 3).DataLabel.Text
End Function

Private Function PtvOf(cht As Chart) As Double
    PtvOf = Util_ToDouble(Registry_Block(cht).Cells(1, WFP_META_PTV).Value)
End Function

' Review Focus 2: os percentuais da imagem do think-cell (custo de -4 para -7 = +75%).
Public Sub Test_Wfp_Arrows_TextsLikeTheImage()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, i As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    cht.Refresh
    AssertEqual "+7%", ArrowLabel(cht, 1): AssertEqual "+75%", ArrowLabel(cht, 2): AssertEqual "-67%", ArrowLabel(cht, 3)
    AssertEqual "+67%", ArrowLabel(cht, 4): AssertEqual "-9%", ArrowLabel(cht, 5)
    For i = 1 To 5
        AssertEqual PctText(SideVal(cht, i, 1, "Val"), SideVal(cht, i, 2, "Val")), ArrowLabel(cht, i), "seta " & i & " pelos valores das barras"
    Next
    AssertEqual "1", Registry_GetOption(cht, "pairdiff"): AssertEqual "0", Registry_GetOption(cht, "pairabs")
    Set s = Chart_FindSeries(cht, "tqq_pd")
    AssertEqual 29, s.Points.Count, "5 pontos por par e uma linha vazia entre pares"
    AssertEqual msoArrowheadTriangle, s.Format.Line.EndArrowheadStyle
    AssertEqual msoShapeOval, s.Points(3).DataLabel.Format.AutoShapeType, "rotulo em elipse"
    AssertEqual xlPrimary, s.AxisGroup
    AssertTrue Not Chart_FindSeries(cht, "tqq_pdroom") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Arrows_NdAndAbs()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Etapa", "S1", "S2")
    ws.Range("A2:C2").Value = Array("Inicio", 10, 10)
    ws.Range("A3:C3").Value = Array("Zero", 0, 3)
    ws.Range("A4:C4").Value = Array("Troca", -2, 4)
    ws.Range("A5:C5").Value = Array("Fim", "e", "e")
    Set d = ChartData_FromRange(ws.Range("A1:C5"))
    Set cht = Wfp_Insert(ws, d).Chart
    cht.Refresh
    AssertEqual "0%", ArrowLabel(cht, 1), "iguais"
    AssertEqual "n/d", ArrowLabel(cht, 2), "a serie 1 vale zero"
    AssertEqual "n/d", ArrowLabel(cht, 3), "sinais opostos"
    AssertEqual "+113%", ArrowLabel(cht, 4), "totais 8 e 17: (17 - 8) / 8 = 112,5% -> 113%"
    Wfp_SetArrowMode cht, True
    cht.Refresh
    AssertEqual "0", ArrowLabel(cht, 1): AssertEqual "+3", ArrowLabel(cht, 2): AssertEqual "+6", ArrowLabel(cht, 3): AssertEqual "+9", ArrowLabel(cht, 4)
    AssertEqual "1", Registry_GetOption(cht, "pairabs")
    CloseTemp wb
End Sub

' Review Focus 3: a seta sai 3 pt acima do degrau e 22 pt acima da barra grande com rotulo (regra de tres no eixo real), sobe 12 pt, a elipse
' fica centrada na linha horizontal, entre as duas barras, dentro do grafico.
Public Sub Test_Wfp_Arrows_Geometry()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, i As Long, r0 As Long, p1 As Point, p2 As Point, p3 As Point, p4 As Point, p5 As Point
    Dim gA As Double, gB As Double, lbl As DataLabel, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    ForceLayout cht                                  ' layout desenhado, calibracao nele, e o desenho final
    Core_SafeRefresh cht
    ForceLayout cht
    Set s = Chart_FindSeries(cht, "tqq_pd"): Set body = BodyOf(cht)
    For i = 1 To 5
        r0 = WFP_ROWS_PER_PAIR * (i - 1)
        Set p1 = s.Points(r0 + 1): Set p2 = s.Points(r0 + 2): Set p3 = s.Points(r0 + 3): Set p4 = s.Points(r0 + 4): Set p5 = s.Points(r0 + 5)
        If CBool(body.Cells(i + 1, Wfp_SideCol(1, "Big")).Value) Then gA = WFP_PT_LBLROOM Else gA = WFP_PT_GAP      ' (rotulos ligados e totais positivos na fixture)
        If CBool(body.Cells(i + 1, Wfp_SideCol(2, "Big")).Value) Then gB = WFP_PT_LBLROOM Else gB = WFP_PT_GAP
        AssertAlmost XOfX(cht, 3 * i - 2), p1.Left, 1, "par " & i & ": sai do centro da barra A"
        AssertAlmost XOfX(cht, 3 * i - 1), p5.Left, 1, "par " & i & ": chega ao centro da barra B"
        AssertAlmost YOf(cht, SideVal(cht, i, 1, "Hi")) - gA, p1.Top, 1.5, "par " & i & ": inicio acima da barra A"
        AssertAlmost YOf(cht, SideVal(cht, i, 2, "Hi")) - gB, p5.Top, 1.5, "par " & i & ": ponta acima da barra B"
        AssertAlmost p2.Top, p4.Top, 0.5, "par " & i & ": trecho horizontal"
        AssertAlmost IIf(p1.Top < p5.Top, p1.Top, p5.Top) - 12, p2.Top, 1.5, "par " & i & ": 12 pt acima do inicio mais alto"
        Set lbl = p3.DataLabel
        AssertAlmost XOfX(cht, 3 * i - 1.5), lbl.Left + lbl.Width / 2, 1.5, "par " & i & ": elipse entre as barras"
        AssertAlmost p3.Top, lbl.Top + lbl.Height / 2, 1.5, "par " & i & ": elipse centrada na linha"
        AssertTrue lbl.Top >= 0, "par " & i & ": elipse dentro do grafico"
    Next
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Arrows_RoomCalibrated_Stable_SmallChart()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ax As Axis, ptv As Double, r As Series, i As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    ForceLayout cht
    Core_SafeRefresh cht
    Set ax = cht.Axes(xlValue, xlPrimary)
    ptv = PtvOf(cht)
    AssertAlmost (ax.MaximumScale - ax.MinimumScale) / cht.PlotArea.InsideHeight, ptv, 0.02 * ptv, "PtV = faixa do eixo / altura interna"
    Core_SafeRefresh cht
    Core_SafeRefresh cht
    AssertAlmost ptv, PtvOf(cht), 0.02 * ptv, "estavel em dois refresh"
    ForceLayout cht
    Set r = Chart_FindSeries(cht, "tqq_pdroom")
    For i = 1 To 5
        AssertTrue r.Points(i).Top >= cht.PlotArea.InsideTop - 0.5, "a sala da elipse " & i & " cabe na escala"
    Next
    cht.Parent.Width = 220: cht.Parent.Height = 140
    Core_SafeRefresh cht
    AssertTrue PtvOf(cht) > 0, "grafico pequeno: a calibracao termina com valor positivo"
    ptv = PtvOf(cht)
    Core_SafeRefresh cht
    AssertAlmost ptv, PtvOf(cht), 0.02 * ptv, "grafico pequeno: estavel no refresh seguinte"
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Arrows_ResizeRecalibrates()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ax As Axis, ptv0 As Double, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    ptv0 = PtvOf(cht)
    cht.Parent.Height = 200
    ForceLayout cht
    Core_SafeRefresh cht
    ForceLayout cht
    Set ax = cht.Axes(xlValue, xlPrimary)
    AssertAlmost (ax.MaximumScale - ax.MinimumScale) / cht.PlotArea.InsideHeight, PtvOf(cht), 0.02 * PtvOf(cht), "recalibrado no tamanho novo"
    AssertTrue PtvOf(cht) > ptv0 * 1.1, "grafico mais baixo: mais valor por ponto (" & ptv0 & " -> " & PtvOf(cht) & ")"
    Set s = Chart_FindSeries(cht, "tqq_pd")
    AssertAlmost YOf(cht, SideVal(cht, 5, 1, "Hi")) - WFP_PT_LBLROOM, s.Points(WFP_ROWS_PER_PAIR * 4 + 1).Top, 1.5, "o afastamento continua 22 pt no total"
    CloseTemp wb
End Sub

Public Sub Test_Wfp_Arrows_ToggleModeAndRefusal()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    d.Source.Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart()
    Core_SetToggle "pairdiff", False
    AssertTrue Chart_FindSeries(cht, "tqq_pd") Is Nothing: AssertTrue Chart_FindSeries(cht, "tqq_pdroom") Is Nothing
    AssertEqual "0", Registry_GetOption(cht, "pairdiff"): AssertFalse Core_GetToggle("pairdiff")
    Core_SetToggle "pairdiff", True
    cht.Refresh
    AssertEqual "+7%", ArrowLabel(cht, 1)
    Core_SetToggle "pairabs", True
    cht.Refresh
    AssertEqual "+2", ArrowLabel(cht, 1): AssertTrue Core_GetToggle("pairabs")
    Core_SetToggle "pairabs", False
    cht.Refresh
    AssertEqual "+7%", ArrowLabel(cht, 1)
    Core_SetOption "legend", "right"
    AssertEqual 2, cht.Legend.LegendEntries.Count, "as auxiliares das setas ficam fora da legenda"
    ws.Range("H1:J1").Value = Array("Ano", "A", "B"): ws.Range("H2:J2").Value = Array(2024, 1, 2): ws.Range("H3:J3").Value = Array(2025, 2, 3)
    ws.Range("H1:J3").Select
    Core_Insert "stacked"
    Core_ClearLastNotice
    Core_SetToggle "pairdiff", True
    AssertEqual "'Setas de variação' não se aplica a este tipo de gráfico.", Core_LastNotice()
    Core_ClearLastNotice
    Core_SetToggle "pairabs", True
    AssertEqual "'Variação em valor' não se aplica a este tipo de gráfico.", Core_LastNotice()
    Teardown wb
End Sub

' Review Focus 5 (setas): sem o ThinQQ, mudar a origem muda o texto da elipse e a posicao da seta (formulas do bloco).
Public Sub Test_Wfp_Arrows_LiveWithoutThinQQ()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, ptv As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    ws.Range("C3").Value = -12
    cht.Refresh
    ptv = PtvOf(cht)
    AssertEqual "+200%", ArrowLabel(cht, 2)
    AssertEqual PctText(11, 5), ArrowLabel(cht, 5), "liquido B = 5"
    AssertEqual 17, SideVal(cht, 3, 2, "Hi"), "B3 agora vai de 17 a 15 (antes 22 a 20)": AssertAlmost SideVal(cht, 3, 2, "Hi") + 3 * ptv, Util_ToDouble(BodyOf(cht).Cells(2 * WFP_ROWS_PER_PAIR + 6, WFP_COL_PY).Value), 0.000001, "ponta do par 3 acompanha o topo novo de B"
    CloseTemp wb
End Sub

' Review Focus 3 (Mesma escala): com o eixo fixado por Deco_SameScale, PtV = (max - min fixos) / altura interna e a seta continua 3 pt acima do
' degrau; Escala automatica volta a calibrar pelo eixo automatico.
Public Sub Test_Wfp_Arrows_SameScale_FixedAxis()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, c1 As Chart, c2 As Chart, coll As Collection, skipped As Long, ax As Axis, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set c1 = Wfp_Insert(ws, d).Chart
    ForceLayout c1
    ws.Range("H1:I1").Value = Array("Ano", "Valor"): ws.Range("H2:I2").Value = Array(2024, 80): ws.Range("H3:I3").Value = Array(2025, 90)
    Set c2 = Stacked_Insert(ws, ChartData_FromRange(ws.Range("H1:I3")), "stacked").Chart
    Set coll = New Collection: coll.Add c1: coll.Add c2
    AssertEqual 2, Deco_SameScale(coll, skipped), "as duas entram (escala livre)"
    Core_SafeRefresh c1
    ForceLayout c1
    Set ax = c1.Axes(xlValue, xlPrimary)
    AssertFalse ax.MaximumScaleIsAuto, "eixo fixo"
    AssertAlmost (ax.MaximumScale - ax.MinimumScale) / c1.PlotArea.InsideHeight, PtvOf(c1), 0.02 * PtvOf(c1), "PtV pelo eixo fixo"
    Set s = Chart_FindSeries(c1, "tqq_pd")
    AssertAlmost YOf(c1, SideVal(c1, 2, 1, "Hi")) - WFP_PT_GAP, s.Points(WFP_ROWS_PER_PAIR + 1).Top, 1.5, "par 2: 3 pt acima do degrau com o eixo fixo"
    AssertFalse Deco_ScaleClipped(c1), "parceiro so acima: a sala das setas cabe no maximo comum"
    AssertTrue Deco_ResetScale(c1)
    Core_SafeRefresh c1
    Set ax = c1.Axes(xlValue, xlPrimary)
    AssertAlmost (ax.MaximumScale - ax.MinimumScale) / c1.PlotArea.InsideHeight, PtvOf(c1), 0.02 * PtvOf(c1), "Escala automatica: recalibrado"
    CloseTemp wb
End Sub

' Um unico refresh deixa nomes e sala coerentes: a sala recalibrada muda a escala e os nomes conferem de novo (medido: em 460 x 160 e 460 x 170
' os nomes ficavam dentro de barras onde nao cabiam ate o refresh seguinte trocar para a legenda). Altura da barra por regra de tres no eixo real.
Public Sub Test_Wfp_Arrows_ResizeKeepsNamesConsistent()
    Dim wb As Workbook, d As clsChartData, cht As Chart, h As Variant, leg As Boolean, barH As Double
    For Each h In Array(160, 170, 200)
        Set wb = NewTempWorkbook(): Set d = Fixture(wb)
        Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
        cht.Parent.Height = CDbl(h)
        ForceLayout cht
        Chart_EnsureRoomAfterChange cht
        leg = cht.HasLegend
        If Not leg Then
            barH = YOf(cht, SideVal(cht, 1, 1, "Lo")) - YOf(cht, SideVal(cht, 1, 1, "Hi"))
            AssertTrue barH >= Labels_EstimateWidth("Segmento 1"), "altura " & h & ": nome dentro de uma barra de " & Format$(barH, "0.0") & " pt"
        End If
        Core_SafeRefresh cht
        AssertEqual leg, cht.HasLegend, "altura " & h & ": o refresh seguinte nao muda a legenda"
        CloseTemp wb
    Next
End Sub

' Falha no meio de ligar as setas (gancho depois de tqq_pd, depois de tqq_pdroom e depois do passo da legenda), com as setas ligadas ou
' desligadas antes e a legenda visivel: nada fica pela metade (nenhuma das duas series, opcao "0", o numero de series sem as setas, legenda com 2
' entradas) e o erro chega ao chamador. Sem o gancho, ligar de novo funciona. O gancho e zerado antes de qualquer afirmacao e no tratamento de erro.
Public Sub Test_Wfp_Arrows_FailMidway_RollsBack()
    Dim wb As Workbook, d As clsChartData, cht As Chart, en As Long, es As String, ed As String, base As Long, site As Variant, prevOn As Variant, tag As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    On Error GoTo Fail
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    Registry_SetOption cht, "legend", "right"
    Wfp_SetLegend cht, "right"
    Wfp_SetArrows cht, False
    base = cht.SeriesCollection.Count
    For Each prevOn In Array(True, False)
        For Each site In Array("pd", "pdroom", "pdlegend")
            tag = site & IIf(prevOn, " (ligadas antes)", " (desligadas antes)")
            Wfp_SetArrows cht, CBool(prevOn)
            Util_FailThinSyncForTest = "wfp-" & site & ":1"
            On Error Resume Next
            Wfp_SetArrows cht, True
            en = Err.Number
            On Error GoTo Fail
            Util_FailThinSyncForTest = ""
            AssertEqual vbObjectError + 9996, en, tag & ": o erro chega ao chamador"
            AssertTrue Chart_FindSeries(cht, "tqq_pd") Is Nothing, tag & ": sem tqq_pd"
            AssertTrue Chart_FindSeries(cht, "tqq_pdroom") Is Nothing, tag & ": sem tqq_pdroom"
            AssertEqual "0", Registry_GetOption(cht, "pairdiff"), tag & ": a opcao reflete o estado real"
            AssertEqual base, cht.SeriesCollection.Count, tag & ": o numero de series sem as setas"
            AssertTrue cht.HasLegend, tag & ": legenda visivel": AssertEqual 2, cht.Legend.LegendEntries.Count, tag & ": legenda com 2 entradas"
            Wfp_SetArrows cht, True
            cht.Refresh
            AssertEqual "1", Registry_GetOption(cht, "pairdiff"), tag & ": ligar de novo funciona"
            AssertEqual base + 2, cht.SeriesCollection.Count, tag
            AssertEqual 2, cht.Legend.LegendEntries.Count, tag & ": as setas ficam fora da legenda"
            AssertEqual "+7%", ArrowLabel(cht, 1), tag
        Next
    Next
    Util_FailThinSyncForTest = ""
    CloseTemp wb
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailThinSyncForTest = ""
    CloseTemp wb
    Err.Raise en, es, ed
End Sub

' Area interna baixa demais para a sala das setas (h < 2 * (22 + 12 + 12) = 92 pt): a calibracao nao age (medido: em 460 x 130 o eixo subia
' de 40 a 120 em um refresh e continuava subindo); PtV e o eixo ficam como estavam em tres refresh seguidos (os afastamentos so encolhem em pontos).
Public Sub Test_Wfp_Arrows_TinyPlot_KeepsPtVAndAxis()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ax As Axis, ptv0 As Double, mx As Double, k As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    ptv0 = PtvOf(cht)
    cht.Parent.Height = 110
    ForceLayout cht
    Core_SafeRefresh cht
    ForceLayout cht
    Set ax = cht.Axes(xlValue, xlPrimary)
    AssertTrue cht.PlotArea.InsideHeight < 2 * (WFP_PT_LBLROOM + WFP_PT_RISE + WFP_PT_ROOM), "area interna abaixo do limite (" & cht.PlotArea.InsideHeight & ")"
    AssertAlmost ptv0, PtvOf(cht), 0.000001, "PtV de antes"
    mx = ax.MaximumScale
    For k = 1 To 3
        Core_SafeRefresh cht
        AssertEqual mx, ax.MaximumScale, "refresh " & k & ": eixo"
        AssertAlmost ptv0, PtvOf(cht), 0.000001, "refresh " & k & ": PtV"
    Next
    CloseTemp wb
End Sub

' Mesma escala em grafico baixo (area interna < 92 pt): com o eixo FIXO a sala das setas nao mexe na escala (sem realimentacao), entao a
' calibracao age e acerta numa passada: PtV = faixa fixa / altura interna, a seta do degrau continua 3 pt acima dele e o eixo fica no fixo.
' (Com o eixo automatico a saida antecipada continua: Test_Wfp_Arrows_TinyPlot_KeepsPtVAndAxis.) Geometria por regra de tres no eixo real.
Public Sub Test_Wfp_Arrows_SameScale_TinyPlot_Calibrates()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, c1 As Chart, c2 As Chart, coll As Collection, skipped As Long, ax As Axis, s As Series
    Dim mx As Double, mn As Double, ptv As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set c1 = Wfp_Insert(ws, d).Chart
    ws.Range("H1:I1").Value = Array("Ano", "Valor"): ws.Range("H2:I2").Value = Array(2024, 80): ws.Range("H3:I3").Value = Array(2025, 90)
    Set c2 = Stacked_Insert(ws, ChartData_FromRange(ws.Range("H1:I3")), "stacked").Chart
    Set coll = New Collection: coll.Add c1: coll.Add c2
    AssertEqual 2, Deco_SameScale(coll, skipped)
    c1.Parent.Height = 110
    ForceLayout c1
    Core_SafeRefresh c1
    ForceLayout c1
    Set ax = c1.Axes(xlValue, xlPrimary)
    AssertTrue c1.PlotArea.InsideHeight < 2 * (WFP_PT_LBLROOM + WFP_PT_RISE + WFP_PT_ROOM), "premissa: area interna abaixo de 92 pt (" & c1.PlotArea.InsideHeight & ")"
    AssertFalse ax.MaximumScaleIsAuto, "eixo fixo"
    mx = ax.MaximumScale: mn = ax.MinimumScale
    ptv = (mx - mn) / c1.PlotArea.InsideHeight
    AssertAlmost ptv, PtvOf(c1), 0.02 * ptv, "PtV pelo eixo fixo e a altura pequena (" & PtvOf(c1) & ")"
    Set s = Chart_FindSeries(c1, "tqq_pd")
    AssertAlmost YOf(c1, SideVal(c1, 2, 1, "Hi")) - WFP_PT_GAP, s.Points(WFP_ROWS_PER_PAIR + 1).Top, 1.5, "par 2: 3 pt acima do degrau"
    Core_SafeRefresh c1
    AssertEqual mx, ax.MaximumScale, "o refresh seguinte nao mexe no eixo"
    AssertAlmost ptv, PtvOf(c1), 0.02 * ptv, "estavel no refresh seguinte"
    CloseTemp wb
End Sub

' Aviso de corte com Mesma escala na comparativa: o grafico encolhido sob a escala fixa recalibra as setas e o aviso das SETAS chega no mesmo
' refresh (a conferencia vem depois da recalibracao); barras que passam do maximo (outro grafico, outro id: o aviso e unico por grafico) dao o
' aviso das barras.
Public Sub Test_Wfp_SameScale_ClipNotice_ArrowsThenBars()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, c1 As Chart, c2 As Chart, c3 As Chart, coll As Collection, skipped As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set c1 = Wfp_Insert(ws, d).Chart
    ws.Range("H1:I1").Value = Array("Ano", "Valor"): ws.Range("H2:I2").Value = Array(2024, 80): ws.Range("H3:I3").Value = Array(2025, 90)
    Set c2 = Kinds_Insert(ws, ChartData_FromRange(ws.Range("H1:I3")), "stacked").Chart
    Set coll = New Collection: coll.Add c1: coll.Add c2
    AssertEqual 2, Deco_SameScale(coll, skipped)
    ForceLayout c1
    Core_ClearLastNotice
    Core_SafeRefresh c1
    AssertEqual 0, Core_NoticeCount(), "tamanho padrao: sem corte, sem aviso"
    c1.Parent.Height = 110
    ForceLayout c1
    Core_ClearLastNotice
    Core_SafeRefresh c1
    AssertEqual "As setas de variação passam do máximo da escala fixa (Mesma escala) e ficam cortadas. Use 'Escala automática' ou desligue as setas de variação.", Core_LastNotice(), "encolhido: o aviso das setas no mesmo refresh"
    AssertFalse Deco_ScaleClipped(c1, True), "as barras cabem"
    Set c3 = Wfp_Insert(ws, d).Chart
    Set coll = New Collection: coll.Add c3: coll.Add c2
    AssertEqual 2, Deco_SameScale(coll, skipped)
    ws.Range("B2").Value = 150
    Core_ClearLastNotice
    Core_SafeRefresh c3
    AssertTrue Deco_ScaleClipped(c3, True), "premissa: a barra do Bruto (150) passa do maximo fixo"
    AssertEqual "Os dados deste gráfico passaram do máximo da escala fixa (Mesma escala) e as barras ficam cortadas. Use 'Mesma escala' de novo ou 'Escala automática'.", Core_LastNotice(), "barras cortadas: o aviso das barras"
    CloseTemp wb
End Sub

' Mesma escala com um parceiro que estende o MINIMO para baixo (cascata de 0 a ~40 e colunas de -40 e 20): a faixa comum quase dobra com o
' mesmo maximo, o valor por ponto dobra e a sala das setas do topo passa do maximo fixo (as barras cabem). Limite conhecido: a escala fica a que
' a Mesma escala calculou; Deco_ScaleClipped conta a sala das setas (tqq_pdroom) e o refresh avisa uma vez. Premissas pelo eixo real.
Public Sub Test_Wfp_SameScale_PartnerExtendsMin_ArrowsClipDetected()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, c1 As Chart, c2 As Chart, coll As Collection, skipped As Long, ax As Axis
    Dim v As Variant, i As Long, mxRoom As Double, values() As Double, tops() As Double, mxTop As Double, autoRange As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set c1 = Wfp_Insert(ws, d).Chart
    ForceLayout c1
    Core_SafeRefresh c1
    Set ax = c1.Axes(xlValue, xlPrimary)
    autoRange = ax.MaximumScale - ax.MinimumScale
    AssertFalse Deco_ScaleClipped(c1), "sem escala fixa nao ha corte"
    ws.Range("H1:I1").Value = Array("Ano", "Valor"): ws.Range("H2:I2").Value = Array(2024, -40): ws.Range("H3:I3").Value = Array(2025, 20)
    Set c2 = Stacked_Insert(ws, ChartData_FromRange(ws.Range("H1:I3")), "stacked").Chart
    Set coll = New Collection: coll.Add c1: coll.Add c2
    AssertEqual 2, Deco_SameScale(coll, skipped)
    ForceLayout c1
    AssertFalse ax.MaximumScaleIsAuto, "eixo fixo"
    AssertTrue ax.MinimumScale < 0 And ax.MaximumScale - ax.MinimumScale >= 1.5 * autoRange, "premissa: o parceiro estendeu o minimo (" & ax.MinimumScale & " a " & ax.MaximumScale & ")"
    v = Chart_FindSeries(c1, "tqq_pdroom").Values
    For i = LBound(v) To UBound(v)
        If Util_ToDouble(v(i)) > mxRoom Then mxRoom = Util_ToDouble(v(i))
    Next
    AssertTrue mxRoom > ax.MaximumScale, "premissa: a sala das setas passa do maximo fixo (" & mxRoom & " > " & ax.MaximumScale & ")"
    Ann_GetCategoryData c1, values, tops
    For i = 1 To UBound(tops)
        If tops(i) > mxTop Then mxTop = tops(i)
    Next
    AssertTrue mxTop <= ax.MaximumScale, "as barras cabem (" & mxTop & ")"
    AssertTrue Deco_ScaleClipped(c1), "as setas cortadas contam como corte"
    AssertFalse Deco_ScaleClipped(c1, True), "so as barras: cabem"
    Core_ClearLastNotice
    Core_SafeRefresh c1
    AssertEqual "As setas de variação passam do máximo da escala fixa (Mesma escala) e ficam cortadas. Use 'Escala automática' ou desligue as setas de variação.", Core_LastNotice(), "o refresh avisa: as setas (nao as barras)"
    Wfp_SetArrows c1, False
    AssertFalse Deco_ScaleClipped(c1), "sem as setas nao ha corte"
    CloseTemp wb
End Sub

' Uma linha com total (marcador "e") de um lado e degrau do outro: comparar o fim acumulado com um degrau nao tem sentido, "n/d" nos dois modos.
' Esperado pela tabela: Inicio 10 | 10 (0%, 0); Passo 2 | 3 ((3 - 2) / 2 = +50%, +1); Sub e | 4 (total x degrau: n/d); Fim e | e (totais 12 e 17:
' (17 - 12) / 12 = 41,7% -> +42%, +5).
Public Sub Test_Wfp_Arrows_TotalVersusStep_Nd()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Etapa", "S1", "S2")
    ws.Range("A2:C2").Value = Array("Inicio", 10, 10)
    ws.Range("A3:C3").Value = Array("Passo", 2, 3)
    ws.Range("A4:C4").Value = Array("Sub", "e", 4)
    ws.Range("A5:C5").Value = Array("Fim", "e", "e")
    Set d = ChartData_FromRange(ws.Range("A1:C5"))
    Set cht = Wfp_Insert(ws, d).Chart
    cht.Refresh
    AssertEqual "0%", ArrowLabel(cht, 1): AssertEqual "+50%", ArrowLabel(cht, 2)
    AssertEqual "n/d", ArrowLabel(cht, 3), "total de um lado, degrau do outro"
    AssertEqual "+42%", ArrowLabel(cht, 4)
    Wfp_SetArrowMode cht, True
    cht.Refresh
    AssertEqual "0", ArrowLabel(cht, 1): AssertEqual "+1", ArrowLabel(cht, 2)
    AssertEqual "n/d", ArrowLabel(cht, 3), "em valor tambem"
    AssertEqual "+5", ArrowLabel(cht, 4)
    ws.Range("B4").Value = 1                         ' os dois lados viram degrau: (4 - 1) / 1 = +300%
    Wfp_SetArrowMode cht, False
    cht.Refresh
    AssertEqual "+300%", ArrowLabel(cht, 3)
    CloseTemp wb
End Sub

' Review Focus 1 (setas): cruza o zero e total negativo (FixtureZero). Esperado pela tabela, sem o bloco: topos (Hi) A = 10, 10, -7, 0 (o total -7
' vai de 0 a -7) e B = 5, 5, 13, 13; afastamento 22 pt nas barras grandes com fim >= 0 (2024 dos dois lados e o total 13 de B) e 3 pt nas demais
' (degraus e o total negativo de A); textos -50%, (-12 + 25) / -25 = -52%, (20 - 8) / 8 = +150% e n/d (totais -7 e 13: sinais opostos).
Public Sub Test_Wfp_Arrows_Geometry_CrossesZero()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, i As Long, r0 As Long, hiA As Variant, hiB As Variant, gA As Variant, gB As Variant
    Dim txt As Variant, yA As Double, yB As Double, lbl As DataLabel
    Set wb = NewTempWorkbook(): Set d = FixtureZero(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    ForceLayout cht
    Core_SafeRefresh cht
    ForceLayout cht
    hiA = Array(10, 10, -7, 0): hiB = Array(5, 5, 13, 13)
    gA = Array(22, 3, 3, 3): gB = Array(22, 3, 3, 22)
    txt = Array("-50%", "-52%", "+150%", "n/d")
    Set s = Chart_FindSeries(cht, "tqq_pd")
    For i = 1 To 4
        r0 = WFP_ROWS_PER_PAIR * (i - 1)
        yA = YOf(cht, CDbl(hiA(i - 1))) - gA(i - 1): yB = YOf(cht, CDbl(hiB(i - 1))) - gB(i - 1)
        AssertAlmost yA, s.Points(r0 + 1).Top, 1.5, "par " & i & ": inicio acima da barra A"
        AssertAlmost yB, s.Points(r0 + 5).Top, 1.5, "par " & i & ": ponta acima da barra B"
        AssertAlmost IIf(yA < yB, yA, yB) - WFP_PT_RISE, s.Points(r0 + 3).Top, 1.5, "par " & i & ": linha 12 pt acima do inicio mais alto"
        Set lbl = s.Points(r0 + 3).DataLabel
        AssertEqual CStr(txt(i - 1)), lbl.Text, "par " & i
        AssertAlmost s.Points(r0 + 3).Top, lbl.Top + lbl.Height / 2, 1.5, "par " & i & ": elipse centrada na linha"
        AssertTrue lbl.Top >= cht.PlotArea.InsideTop - 0.5, "par " & i & ": elipse dentro da area"
    Next
    AssertAlmost YOf(cht, 0) - WFP_PT_GAP, s.Points(WFP_ROWS_PER_PAIR * 3 + 1).Top, 1.5, "total negativo: a seta sai 3 pt acima do zero"
    CloseTemp wb
End Sub

Private Function LblText(cht As Chart, seriesName As String, k As Long) As String
    LblText = Chart_FindSeries(cht, seriesName).Points(k).DataLabel.Text
End Function

Public Sub Test_Wfp_Labels_StepsCenterTotalsAbove()
    Dim wb As Workbook, d As clsChartData, cht As Chart, la As Series, lt As Series, lbl As DataLabel, p As Point
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    cht.Refresh
    AssertEqual "1", Registry_GetOption(cht, "labels")
    AssertEqual Format$(-4, "#,##0"), LblText(cht, "tqq_lbl_a", 4), "degrau A2 assinado"
    AssertEqual Format$(-7, "#,##0"), LblText(cht, "tqq_lbl_b", 5), "degrau B2"
    AssertEqual Format$(-2, "#,##0"), LblText(cht, "tqq_lbl_b", 8), "degrau B3 de 2 visivel no tamanho padrao"
    AssertEqual "", LblText(cht, "tqq_lbl_a", 1), "barra grande: sem rotulo de degrau"
    AssertEqual "", LblText(cht, "tqq_lbl_a", 3), "vao"
    AssertEqual Format$(27, "#,##0"), LblText(cht, "tqq_lbltot", 1), "barra inicial: so o valor"
    AssertEqual Format$(11, "#,##0"), LblText(cht, "tqq_lbltot", 13), "460 x 280: os totais do par se tocariam (41 pt numa fatia de 31): so o valor"
    cht.Parent.Width = 720                           ' fatia de ~50 pt: o percentual cabe
    ForceLayout cht
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual Format$(11, "#,##0") & " (" & Format$(41, "0") & "%)", LblText(cht, "tqq_lbltot", 13), "total A: valor e % do 1o total (11 / 27)"
    AssertEqual Format$(10, "#,##0") & " (" & Format$(34, "0") & "%)", LblText(cht, "tqq_lbltot", 14), "total B (10 / 29)"
    Set la = Chart_FindSeries(cht, "tqq_lbl_a"): Set lt = Chart_FindSeries(cht, "tqq_lbltot")
    AssertEqual xlLabelPositionCenter, la.Points(4).DataLabel.Position
    AssertEqual xlLabelPositionAbove, lt.Points(13).DataLabel.Position
    AssertTrue lt.Points(13).DataLabel.Font.Bold
    AssertEqual RGB(255, 255, 255), la.Points(4).DataLabel.Font.Color, "contraste na barra escura"
    AssertEqual RGB(64, 64, 64), Chart_FindSeries(cht, "tqq_lbl_b").Points(5).DataLabel.Font.Color, "contraste na barra clara"
    ForceLayout cht
    Set lbl = la.Points(4).DataLabel
    AssertTrue lbl.Top + lbl.Height / 2 > YOf(cht, 27) And lbl.Top + lbl.Height / 2 < YOf(cht, 23), "rotulo do degrau dentro da barra (de 27 a 23)"
    Set lbl = lt.Points(13).DataLabel: Set p = cht.SeriesCollection(2).Points(13)
    AssertTrue Chart_FindSeries(cht, "tqq_pd").Points(WFP_ROWS_PER_PAIR * 4 + 1).Top <= lbl.Top - 1, "a seta do total comeca acima do rotulo do total"
    AssertTrue lbl.Top + lbl.Height <= p.Top + 1, "rotulo do total acima da barra"
    CloseTemp wb
End Sub

' Review Focus 1 (rotulos): total negativo com o rotulo ABAIXO da barra.
Public Sub Test_Wfp_Labels_NegativeTotalBelow()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lt As Series, lbl As DataLabel, p As Point
    Set wb = NewTempWorkbook(): Set d = FixtureZero(wb)
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    cht.Refresh
    Set lt = Chart_FindSeries(cht, "tqq_lbltot")
    AssertEqual Format$(-7, "#,##0") & " (" & Format$(-70, "0") & "%)", LblText(cht, "tqq_lbltot", 10), "total A -7 de um inicio de 10"
    AssertEqual xlLabelPositionBelow, lt.Points(10).DataLabel.Position, "total negativo: abaixo"
    AssertEqual xlLabelPositionAbove, lt.Points(11).DataLabel.Position, "total positivo: acima"
    ForceLayout cht
    Set lbl = lt.Points(10).DataLabel: Set p = cht.SeriesCollection(3).Points(10)
    AssertTrue lbl.Top >= p.Top + p.Height - 1, "abaixo da barra negativa"
    CloseTemp wb
End Sub

' Degrau fino demais: texto vazio (LblOn = 0) e volta quando cabe; rotulo apagado pelo usuario continua apagado.
Public Sub Test_Wfp_Labels_ThinTextHidden_UserDeletedStays()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, la As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    ws.Range("B4").Value = -0.2
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual 0, Util_ToDouble(BodyOf(cht).Cells(8, WFP_COL_LBLON).Value), "degrau A3 de 0,2: fino demais"
    AssertEqual "", LblText(cht, "tqq_lbl_a", 7)
    ws.Range("B4").Value = -6
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual Format$(-6, "#,##0"), LblText(cht, "tqq_lbl_a", 7), "voltou a caber"
    Set la = Chart_FindSeries(cht, "tqq_lbl_a")
    la.Points(4).HasDataLabel = False                ' o usuario apagou o rotulo do degrau A2
    Core_SafeRefresh cht
    AssertFalse la.Points(4).HasDataLabel, "o refresh nao recria rotulo apagado pelo usuario"
    Teardown wb
End Sub

Public Sub Test_Wfp_Totals_PctToggle_And_NumFmt()
    Dim wb As Workbook, d As clsChartData, cht As Chart, dec As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart()
    cht.Parent.Width = 800                           ' fatia de ~55 pt: "11,0 (41%)" (49 pt estimados) cabe sem tocar o vizinho (margem folgada)
    ForceLayout cht
    Core_SafeRefresh cht
    AssertTrue Wfp_TotalPctOn(cht): AssertTrue Core_GetToggle("totals")
    Core_SetToggle "totals", False
    cht.Refresh
    AssertEqual Format$(11, "#,##0"), LblText(cht, "tqq_lbltot", 13), "sem o percentual"
    AssertFalse Wfp_TotalPctOn(cht): AssertEqual "0", Registry_GetOption(cht, "totals")
    Core_SetToggle "totals", True
    Core_SetOption "numfmt", "dec1"
    cht.Refresh
    dec = Application.International(xlDecimalSeparator)
    AssertEqual "-4" & dec & "0", LblText(cht, "tqq_lbl_a", 4), "degrau com 1 casa"
    AssertEqual "11" & dec & "0 (41%)", LblText(cht, "tqq_lbltot", 13), "total com 1 casa e o percentual inteiro"
    Core_SetToggle "pairabs", True
    cht.Refresh
    AssertEqual "+2" & dec & "0", Chart_FindSeries(cht, "tqq_pd").Points(3).DataLabel.Text, "variacao em valor no formato do grafico"
    Core_SetOption "numfmt", "pct0"
    AssertEqual " p.p.", CStr(Registry_Block(cht).Cells(1, WFP_META_SUFD).Value), "diferenca de percentuais em pontos percentuais"
    Teardown wb
End Sub

Public Sub Test_Wfp_Labels_Toggle_MovesArrowStart()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, k As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Select
    Core_Insert "waterfallpair"
    Set cht = Core_ActiveChart()
    k = WFP_ROWS_PER_PAIR * 4 + 1                    ' P1 do par 5 (Liquido A, barra grande)
    Core_SetToggle "labels", False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl_a") Is Nothing: AssertTrue Chart_FindSeries(cht, "tqq_lbltot") Is Nothing
    AssertEqual 0, Util_ToDouble(Registry_Block(cht).Cells(1, WFP_META_LBL).Value)
    ForceLayout cht
    Set s = Chart_FindSeries(cht, "tqq_pd")
    AssertAlmost YOf(cht, SideVal(cht, 5, 1, "Hi")) - 3, s.Points(k).Top, 1.5, "sem rotulos a seta sai 3 pt acima do total"
    Core_SetToggle "labels", True
    ForceLayout cht
    AssertAlmost YOf(cht, SideVal(cht, 5, 1, "Hi")) - WFP_PT_LBLROOM, s.Points(k).Top, 1.5, "com rotulos, 22 pt (acima do rotulo do total)"
    AssertTrue Not Chart_FindSeries(cht, "tqq_lbl_b") Is Nothing
    Core_SetOption "legend", "right"
    AssertEqual 2, cht.Legend.LegendEntries.Count, "rotulos fora da legenda"
    Teardown wb
End Sub

' Um degrau que vira subtotal e um total que vira degrau DEPOIS de inserir: os pontos que eram NA() na insercao aparecem ligados, com o texto,
' a posicao e a cor certos (o refresh nao religa rotulos).
Public Sub Test_Wfp_Labels_MarkerChange_PointsThatWereNA()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    ws.Range("C5").Value = "e"                       ' Componente 3 de B vira subtotal (fatia 11): 29 - 7 - 2 = 20
    ws.Range("B6").Value = -5                        ' Liquido de A vira degrau (fatia 13)
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual Format$(20, "#,##0") & " (" & Format$(69, "0") & "%)", LblText(cht, "tqq_lbltot", 11), "subtotal novo de B (20 / 29)"
    AssertEqual xlLabelPositionAbove, Chart_FindSeries(cht, "tqq_lbltot").Points(11).DataLabel.Position
    AssertEqual Format$(-5, "#,##0"), LblText(cht, "tqq_lbl_a", 13), "degrau novo de A"
    AssertEqual xlLabelPositionCenter, Chart_FindSeries(cht, "tqq_lbl_a").Points(13).DataLabel.Position
    AssertEqual RGB(255, 255, 255), Chart_FindSeries(cht, "tqq_lbl_a").Points(13).DataLabel.Font.Color
    CloseTemp wb
End Sub

' Falha no meio de ligar os rotulos (gancho depois de tqq_lbl_a, de tqq_lbl_b, de tqq_lbltot e depois do passo da legenda), com os rotulos
' ligados ou desligados antes e a legenda visivel: nada fica pela metade (nenhuma das tres series, opcao "0", meta Lbl 0, o numero de series sem
' os rotulos, legenda com 2 entradas) e o erro chega ao chamador. Sem o gancho, ligar de novo funciona. O gancho e zerado antes de qualquer
' afirmacao e no tratamento de erro.
Public Sub Test_Wfp_Labels_FailMidway_RollsBack()
    Dim wb As Workbook, d As clsChartData, cht As Chart, en As Long, es As String, ed As String, base As Long, site As Variant, prevOn As Variant, tag As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    On Error GoTo Fail
    Set cht = Wfp_Insert(d.Source.Worksheet, d).Chart
    Registry_SetOption cht, "legend", "right"
    Wfp_SetLegend cht, "right"
    cht.Parent.Width = 760                           ' o percentual cabe ao lado do vizinho com a legenda a direita
    ForceLayout cht
    Wfp_SetLabels cht, False
    base = cht.SeriesCollection.Count
    For Each prevOn In Array(True, False)
        For Each site In Array("lbl:a", "lbl:b", "lbl:tot", "lbllegend:1")
            tag = site & IIf(prevOn, " (ligados antes)", " (desligados antes)")
            Wfp_SetLabels cht, CBool(prevOn)
            Util_FailThinSyncForTest = "wfp-" & site
            On Error Resume Next
            Wfp_SetLabels cht, True
            en = Err.Number
            On Error GoTo Fail
            Util_FailThinSyncForTest = ""
            AssertEqual vbObjectError + 9996, en, tag & ": o erro chega ao chamador"
            AssertTrue Chart_FindSeries(cht, "tqq_lbl_a") Is Nothing, tag & ": sem tqq_lbl_a"
            AssertTrue Chart_FindSeries(cht, "tqq_lbl_b") Is Nothing, tag & ": sem tqq_lbl_b"
            AssertTrue Chart_FindSeries(cht, "tqq_lbltot") Is Nothing, tag & ": sem tqq_lbltot"
            AssertEqual "0", Registry_GetOption(cht, "labels"), tag & ": a opcao reflete o estado real"
            AssertEqual 0, Util_ToDouble(Registry_Block(cht).Cells(1, WFP_META_LBL).Value), tag & ": meta Lbl 0 (as setas descem)"
            AssertEqual base, cht.SeriesCollection.Count, tag & ": o numero de series sem os rotulos"
            AssertTrue cht.HasLegend, tag & ": legenda visivel": AssertEqual 2, cht.Legend.LegendEntries.Count, tag & ": legenda com 2 entradas"
            Wfp_SetLabels cht, True
            cht.Refresh
            AssertEqual "1", Registry_GetOption(cht, "labels"), tag & ": ligar de novo funciona"
            AssertEqual base + 3, cht.SeriesCollection.Count, tag
            AssertEqual 2, cht.Legend.LegendEntries.Count, tag & ": os rotulos ficam fora da legenda"
            AssertEqual Format$(11, "#,##0") & " (" & Format$(41, "0") & "%)", LblText(cht, "tqq_lbltot", 13), tag
        Next
    Next
    Util_FailThinSyncForTest = ""
    CloseTemp wb
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailThinSyncForTest = ""
    CloseTemp wb
    Err.Raise en, es, ed
End Sub

' Caixas REAIS (depois do ForceLayout) dos rotulos de total das fatias k (A) e k + 1 (B) da mesma etapa: a de B comeca depois do fim da de A.
Private Sub AssertTotalBoxesApart(cht As Chart, k As Long, tag As String)
    Dim lt As Series, a As DataLabel, b As DataLabel
    Set lt = Chart_FindSeries(cht, "tqq_lbltot"): Set a = lt.Points(k).DataLabel: Set b = lt.Points(k + 1).DataLabel
    AssertTrue b.Left >= a.Left + a.Width, tag & ": as caixas nao se sobrepoem (A termina em " & Format$(a.Left + a.Width, "0.0") & ", B comeca em " & Format$(b.Left, "0.0") & ")"
End Sub

' Totais vizinhos (os dois lados sao barras grandes) que se tocariam: so o valor nas duas fatias (LblOn = 0), o percentual volta quando cabe.
' Medido (sonda): fixture a 300 x 220 (o teste usa 320 x 220, com folga) e 380 x 240, caixas "11 (41%)" / "10 (34%)" sobrepostas 18,6 e 12,9 pt; 4 digitos a 460 x 280
' ("1.234 (62%)" / "1.230 (62%)"), 18,7 pt. O texto continua vivo sem o ThinQQ (so a coluna LblOn fica como esta).
Public Sub Test_Wfp_Labels_TotalsCollide_ValueOnly()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, lt As Series, a As DataLabel, b As DataLabel
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    cht.Parent.Width = 320: cht.Parent.Height = 220
    ForceLayout cht
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual 0, Util_ToDouble(BodyOf(cht).Cells(14, WFP_COL_LBLON).Value), "fatia 13 (total A): so o valor"
    AssertEqual 0, Util_ToDouble(BodyOf(cht).Cells(15, WFP_COL_LBLON).Value), "fatia 14 (total B): so o valor"
    AssertEqual Format$(11, "#,##0"), LblText(cht, "tqq_lbltot", 13), "320 x 220: total A sem o percentual"
    AssertEqual Format$(10, "#,##0"), LblText(cht, "tqq_lbltot", 14), "320 x 220: total B sem o percentual"
    AssertEqual Format$(27, "#,##0"), LblText(cht, "tqq_lbltot", 1), "inicio: cabe (16 pt)"
    ForceLayout cht
    AssertTotalBoxesApart cht, 1, "320 x 220 (so o valor), inicio"
    AssertTotalBoxesApart cht, 13, "320 x 220 (so o valor), liquido"
    ws.Range("B2").Value = 30                        ' sem o ThinQQ (so o calculo): o texto acompanha
    cht.Refresh
    AssertEqual Format$(14, "#,##0"), LblText(cht, "tqq_lbltot", 13), "vivo sem o ThinQQ"
    ws.Range("B2").Value = 27
    cht.Parent.Width = 720: cht.Parent.Height = 280
    ForceLayout cht
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual 1, Util_ToDouble(BodyOf(cht).Cells(14, WFP_COL_LBLON).Value)
    AssertEqual Format$(11, "#,##0") & " (" & Format$(41, "0") & "%)", LblText(cht, "tqq_lbltot", 13), "720 x 280: o percentual volta"
    AssertEqual Format$(10, "#,##0") & " (" & Format$(34, "0") & "%)", LblText(cht, "tqq_lbltot", 14)
    ForceLayout cht
    Set lt = Chart_FindSeries(cht, "tqq_lbltot"): Set a = lt.Points(13).DataLabel: Set b = lt.Points(14).DataLabel
    AssertTrue b.Left >= a.Left + a.Width, "720 x 280: as caixas nao se sobrepoem (" & a.Left + a.Width & " x " & b.Left & ")"
    ws.Range("B2:C2").Value = Array(2000, 1990): ws.Range("B3:C3").Value = Array(-300, -360)
    ws.Range("B4:C4").Value = Array(-266, -200): ws.Range("B5:C5").Value = Array(-200, -200)
    cht.Parent.Width = 460
    ForceLayout cht
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual Format$(1234, "#,##0"), LblText(cht, "tqq_lbltot", 13), "4 digitos a 460 x 280: totais 1.234 e 1.230 lado a lado, so o valor"
    AssertEqual Format$(1230, "#,##0"), LblText(cht, "tqq_lbltot", 14)
    ForceLayout cht
    AssertTotalBoxesApart cht, 1, "4 digitos a 460 x 280 (so o valor), inicio"
    AssertTotalBoxesApart cht, 13, "4 digitos a 460 x 280 (so o valor), liquido"
    Teardown wb
End Sub

' 1o total = 0: o percentual e omitido; 1o total negativo: -10 -> 5 da "(-50%)" (100 * 5 / -10).
Public Sub Test_Wfp_Totals_FirstTotalZeroOrNegative()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Etapa", "S1", "S2")
    ws.Range("A2:C2").Value = Array("Inicio", 0, -10)
    ws.Range("A3:C3").Value = Array("Passo", 5, 15)
    ws.Range("A4:C4").Value = Array("Fim", "e", "e")
    Set d = ChartData_FromRange(ws.Range("A1:C4"))
    Set cht = Wfp_Insert(ws, d).Chart
    ForceLayout cht
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual Format$(5, "#,##0"), LblText(cht, "tqq_lbltot", 7), "1o total zero: sem o percentual"
    AssertEqual Format$(5, "#,##0") & " (" & Format$(-50, "0") & "%)", LblText(cht, "tqq_lbltot", 8), "1o total -10: 5 / -10 = -50%"
    AssertEqual Format$(-10, "#,##0"), LblText(cht, "tqq_lbltot", 2), "inicio negativo: so o valor"
    CloseTemp wb
End Sub

' Degrau escondido so pela largura: alto o bastante (regra de tres no eixo real >= 10 pt), mas "-4.000" (32 pt estimados) nao cabe na fatia
' de ~27 pt a 400 pt de largura; volta a 720.
Public Sub Test_Wfp_Labels_StepHiddenByWidthOnly()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, ax As Axis, slotPt As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("B2:C2").Value = Array(27000, 29000): ws.Range("B3:C3").Value = Array(-4000, -7000)
    ws.Range("B4:C4").Value = Array(-6000, -2000): ws.Range("B5:C5").Value = Array(-6000, -10000)
    Set cht = Wfp_Insert(ws, d).Chart
    cht.Parent.Width = 400
    ForceLayout cht
    Core_SafeRefresh cht
    cht.Refresh
    Set ax = cht.Axes(xlValue, xlPrimary)
    slotPt = cht.PlotArea.InsideWidth / SlotN(cht)
    AssertTrue 4000 * cht.PlotArea.InsideHeight / (ax.MaximumScale - ax.MinimumScale) >= 10, "alto o bastante: a regra da altura nao age"
    AssertTrue Labels_EstimateWidth(Format$(-4000, "#,##0")) > slotPt, "o texto e mais largo que a fatia (" & slotPt & " pt)"
    AssertEqual 0, Util_ToDouble(BodyOf(cht).Cells(5, WFP_COL_LBLON).Value), "degrau A2 escondido pela largura"
    AssertEqual "", LblText(cht, "tqq_lbl_a", 4)
    cht.Parent.Width = 720
    ForceLayout cht
    Core_SafeRefresh cht
    cht.Refresh
    AssertEqual Format$(-4000, "#,##0"), LblText(cht, "tqq_lbl_a", 4), "720: cabe e volta"
    CloseTemp wb
End Sub

' Rotulo de um total negativo apagado pelo usuario continua apagado no refresh (o acerto de posicao pula ponto sem rotulo), mesmo quando o
' total vira positivo (a posicao mudaria para Acima).
Public Sub Test_Wfp_Labels_UserDeletedNegativeTotalStays()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, lt As Series
    Set wb = NewTempWorkbook(): Set d = FixtureZero(wb): Set ws = d.Source.Worksheet
    Set cht = Wfp_Insert(ws, d).Chart
    Set lt = Chart_FindSeries(cht, "tqq_lbltot")
    AssertEqual xlLabelPositionBelow, lt.Points(10).DataLabel.Position, "total -7: abaixo"
    lt.Points(10).HasDataLabel = False               ' o usuario apagou o rotulo do total negativo de A
    Core_SafeRefresh cht
    AssertFalse lt.Points(10).HasDataLabel, "o refresh nao recria"
    ws.Range("B3").Value = -2                        ' total de A: 10 - 2 + 8 = 16 (positivo)
    Core_SafeRefresh cht
    AssertFalse lt.Points(10).HasDataLabel, "continua apagado com o total positivo"
    AssertTrue lt.Points(11).HasDataLabel: AssertEqual xlLabelPositionAbove, lt.Points(11).DataLabel.Position, "o vizinho continua"
    Teardown wb
End Sub
