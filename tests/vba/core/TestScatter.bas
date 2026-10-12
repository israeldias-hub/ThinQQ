Attribute VB_Name = "TestScatter"
Option Explicit

Private Function Fixture(wb As Workbook, Optional withGroup As Boolean = True, Optional bubble As Boolean = False) As clsChartData
    Dim ws As Worksheet, lastCol As String
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Produto", "Preço", "Margem")
    ws.Range("A2:C2").Value = Array("P1", 10, 0.2): ws.Range("A3:C3").Value = Array("P2", 20, 0.35)
    ws.Range("A4:C4").Value = Array("P3", 30, 0.25): ws.Range("A5:C5").Value = Array("P4", 40, 0.5)
    lastCol = "C"
    If bubble Then
        ws.Range("D1").Value = "Volume": ws.Range("D2:D5").Value = Application.Transpose(Array(50, 120, 80, 20)): lastCol = "D"
    End If
    If withGroup Then
        ws.Range(IIf(bubble, "E1", "D1")).Value = "Grupo"
        ws.Range(IIf(bubble, "E2:E5", "D2:D5")).Value = Application.Transpose(Array("A", "B", "A", "B")): lastCol = IIf(bubble, "E", "D")
    End If
    Set Fixture = ChartData_FromRange(ws.Range("A1:" & lastCol & "5"))
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

Private Function XOf(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlCategory, xlPrimary)
    XOf = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (v - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
End Function

Private Function YOf(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue, xlPrimary)
    YOf = cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight * (ax.MaximumScale - v) / (ax.MaximumScale - ax.MinimumScale)
End Function

Public Sub Test_Scatter_Insert_BlockSeriesAxesLabels()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range, body As Range, sa As Series, sb As Series, v As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Scatter_Insert(d.Source.Worksheet, d, "scatter").Chart
    AssertEqual "scatter", Registry_ChartType(cht)
    AssertEqual xlXYScatter, cht.ChartType
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 11, blk.Columns.Count, "5 + 3 colunas por grupo (2 grupos)": AssertEqual 2, Scatter_GroupCount(blk)
    AssertEqual "Xg1", body.Cells(1, Scatter_XgCol(blk, 1)).Value: AssertEqual "Sg2", body.Cells(1, Scatter_SgCol(blk, 2)).Value
    AssertEqual 10, body.Cells(2, SC_COL_X).Value: AssertEqual 0.2, body.Cells(2, SC_COL_Y).Value: AssertEqual "A", body.Cells(2, SC_COL_GROUP).Value
    AssertEqual 10, body.Cells(2, Scatter_XgCol(blk, 1)).Value, "P1 e do grupo A"
    AssertTrue IsError(body.Cells(2, Scatter_XgCol(blk, 2)).Value), "P1 nao esta no grupo B"
    AssertEqual 20, body.Cells(3, Scatter_XgCol(blk, 2)).Value
    AssertAlmost 25, CDbl(blk.Cells(1, 5).Value), 0.000001, "AvgX": AssertAlmost 0.325, CDbl(blk.Cells(1, 6).Value), 0.000001, "AvgY"
    AssertEqual 2, cht.SeriesCollection.Count
    Set sa = cht.SeriesCollection(1): Set sb = cht.SeriesCollection(2)
    AssertEqual "A", sa.Name: AssertEqual "B", sb.Name
    AssertTrue InStr(sa.Formula, CALC_SHEET) > 0, "series apontam para o bloco"
    AssertEqual xlNotPlotted, cht.DisplayBlanksAs
    v = sa.Values
    AssertTrue IsError(v(2)) Or IsEmpty(v(2)), "ponto do outro grupo nao plotado"
    AssertEqual "P1", sa.Points(1).DataLabel.Text: AssertEqual "P2", sb.Points(2).DataLabel.Text
    AssertEqual xlLabelPositionRight, sa.Points(1).DataLabel.Position
    AssertTrue cht.Axes(xlCategory).HasTitle And cht.Axes(xlValue).HasTitle
    AssertEqual "Preço", cht.Axes(xlCategory).AxisTitle.Text: AssertEqual "Margem", cht.Axes(xlValue).AxisTitle.Text
    AssertEqual xlTickLabelPositionLow, cht.Axes(xlCategory).TickLabelPosition
    AssertEqual xlTickLabelPositionLow, cht.Axes(xlValue).TickLabelPosition
    AssertFalse cht.Axes(xlValue).HasMajorGridlines
    AssertTrue cht.HasLegend: AssertEqual 2, cht.Legend.LegendEntries.Count
    AssertEqual "0", Registry_GetOption(cht, "trend"): AssertEqual "0", Registry_GetOption(cht, "quad"): AssertEqual "4", Registry_GetOption(cht, "n")
    d.Source.Worksheet.Range("B1").Value = "Preço (R$)"
    AssertEqual "Preço (R$)", cht.Axes(xlCategory).AxisTitle.Text, "titulo do eixo ligado ao cabecalho"
    CloseTemp wb
End Sub

Public Sub Test_Scatter_Geometry_PointsAndLabels()
    Dim wb As Workbook, d As clsChartData, cht As Chart, p As Point, lbl As DataLabel
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Scatter_Insert(d.Source.Worksheet, d, "scatter").Chart
    ForceLayout cht
    Set p = cht.SeriesCollection(1).Points(1)                   ' P1 (10; 0,2)
    AssertAlmost XOf(cht, 10), p.Left + p.Width / 2, 2, "centro do marcador em x(10)"
    AssertAlmost YOf(cht, 0.2), p.Top + p.Height / 2, 2, "centro do marcador em y(0,2)"
    Set lbl = p.DataLabel
    AssertTrue lbl.Left >= p.Left + p.Width / 2, "nome a direita do ponto"
    AssertTrue Abs((lbl.Top + lbl.Height / 2) - (p.Top + p.Height / 2)) < 8, "nome na altura do ponto"
    Set p = cht.SeriesCollection(2).Points(4)                   ' P4 (40; 0,5)
    AssertAlmost XOf(cht, 40), p.Left + p.Width / 2, 2
    AssertAlmost YOf(cht, 0.5), p.Top + p.Height / 2, 2
    CloseTemp wb
End Sub

' Review Focus 2: um ponto troca de grupo e outro ganha texto em X.
Public Sub Test_Scatter_GroupChange_MovesPointAndLabel()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, sa As Series, sb As Series, va As Variant, vb As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Scatter_Insert(ws, d, "scatter").Chart
    Set sa = cht.SeriesCollection(1): Set sb = cht.SeriesCollection(2)
    ws.Range("D2").Value = "B"                                  ' P1: A -> B
    ws.Range("B4").Value = "x"                                  ' P3: X invalido
    cht.Refresh
    va = sa.Values: vb = sb.Values
    AssertTrue IsError(va(1)) Or IsEmpty(va(1)), "P1 saiu da serie A"
    AssertAlmost 0.2, CDbl(vb(1)), 0.000001, "P1 entrou na serie B"
    AssertEqual "P1", sb.Points(1).DataLabel.Text, "o nome acompanha o ponto na serie B"
    AssertTrue IsError(va(3)) Or IsEmpty(va(3)), "P3 com X invalido nao e plotado"
    AssertTrue IsError(Block_Body(Registry_Block(cht)).Cells(4, SC_COL_X).Value), "X invalido vira NA() no bloco"
    ForceLayout cht
    AssertTrue sb.Points(1).DataLabel.Left > 0, "rotulo de P1 na serie B tem posicao"
    CloseTemp wb
End Sub

Public Sub Test_Scatter_Trend_And_Quadrants()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, blk As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Scatter_Insert(ws, d, "scatter").Chart
    Set blk = Registry_Block(cht)
    Scatter_SetTrend cht, True
    AssertEqual 1, cht.SeriesCollection(1).Trendlines.Count: AssertEqual 1, cht.SeriesCollection(2).Trendlines.Count
    AssertEqual xlLinear, cht.SeriesCollection(1).Trendlines(1).Type
    AssertEqual 2, cht.Legend.LegendEntries.Count, "entradas das tendencias fora da legenda"
    AssertEqual "1", Registry_GetOption(cht, "trend")
    Scatter_SetTrend cht, False
    AssertEqual 0, cht.SeriesCollection(1).Trendlines.Count: AssertEqual "0", Registry_GetOption(cht, "trend")
    Scatter_SetQuad cht, True
    AssertEqual xlAxisCrossesCustom, cht.Axes(xlCategory).Crosses: AssertEqual xlAxisCrossesCustom, cht.Axes(xlValue).Crosses
    AssertAlmost CDbl(blk.Cells(1, 5).Value), CDbl(cht.Axes(xlCategory).CrossesAt), 0.000001, "linha vertical na media de X"
    AssertAlmost CDbl(blk.Cells(1, 6).Value), CDbl(cht.Axes(xlValue).CrossesAt), 0.000001, "linha horizontal na media de Y"
    AssertEqual xlTickLabelPositionLow, cht.Axes(xlValue).TickLabelPosition, "rotulos ficam na borda"
    ws.Range("B5").Value = 80                                   ' media de X passa de 25 para 35
    Scatter_RefreshAfterDataChange cht
    AssertAlmost 35, CDbl(cht.Axes(xlCategory).CrossesAt), 0.000001, "quadrante acompanha os dados no refresh"
    Scatter_SetQuad cht, False
    AssertEqual xlAxisCrossesAutomatic, cht.Axes(xlCategory).Crosses: AssertEqual "0", Registry_GetOption(cht, "quad")
    CloseTemp wb
End Sub

Public Sub Test_Bubble_Insert_SizesNegativeAndLabels()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, blk As Range, body As Range, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb, True, True): Set ws = d.Source.Worksheet
    ws.Range("D5").Value = -20                                  ' tamanho negativo: fora
    Set cht = Scatter_Insert(ws, d, "bubble").Chart
    AssertEqual "bubble", Registry_ChartType(cht): AssertEqual xlBubble, cht.ChartType
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 50, body.Cells(2, SC_COL_SIZE).Value
    AssertTrue IsError(body.Cells(5, SC_COL_SIZE).Value), "tamanho negativo vira NA()"
    AssertTrue IsError(body.Cells(5, Scatter_XgCol(blk, 2)).Value), "ponto sem tamanho nao e plotado"
    Set s = cht.SeriesCollection(1)
    AssertTrue InStr(s.BubbleSizes, CALC_SHEET) > 0, "tamanhos no bloco"
    AssertEqual 60, cht.ChartGroups(1).BubbleScale
    AssertFalse cht.ChartGroups(1).ShowNegativeBubbles
    AssertEqual xlLabelPositionCenter, s.Points(1).DataLabel.Position
    AssertEqual "P1", s.Points(1).DataLabel.Text
    AssertEqual "Tam", body.Cells(1, SC_COL_SIZE).Value, "cabecalho fixo do bloco (nao o da tabela)"
    CloseTemp wb
End Sub

Public Sub Test_Scatter_Refusals_Rollback_NoGroup()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, en As Long, nm As Name, i As Long, k As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb, False): Set ws = d.Source.Worksheet
    ws.Range("A1:B5").Select
    Core_ClearLastNotice
    Core_Insert "scatter"
    AssertEqual "Dispersão: a tabela precisa de nome, X e Y (e, nas bolhas, tamanho).", Core_LastNotice()
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    ws.Range("A1:C5").Select
    Core_ClearLastNotice
    Core_Insert "bubble"
    AssertEqual "Dispersão: a tabela precisa de nome, X e Y (e, nas bolhas, tamanho).", Core_LastNotice()
    AssertEqual 0, ws.ChartObjects.Count
    Set cht = Scatter_Insert(ws, d, "scatter").Chart            ' sem coluna de grupo: uma serie com o nome do cabecalho de Y
    AssertEqual 1, cht.SeriesCollection.Count: AssertEqual "Margem", cht.SeriesCollection(1).Name
    AssertEqual 8, Registry_Block(cht).Columns.Count
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    ws.Range("D1").Value = "Grupo"
    For i = 2 To 5: ws.Range("D" & i).Value = "G" & i: Next
    ws.Range("A6:D8").Value = Array("P5", 1, 1, "G6"): ws.Range("A7:D7").Value = Array("P6", 2, 2, "G7"): ws.Range("A8:D8").Value = Array("P7", 3, 3, "G8")
    Set d = ChartData_FromRange(ws.Range("A1:D8"))
    On Error Resume Next
    Scatter_Insert ws, d, "scatter"
    en = Err.Number
    On Error GoTo 0
    AssertEqual vbObjectError + 6013, en, "7 grupos"
    AssertEqual 0, ws.ChartObjects.Count: AssertEqual 0, TqqNameCount(wb)
    Set d = ChartData_FromRange(ws.Range("A1:D5"))
    For Each k In Array(9999, 9998)
        If k = 9999 Then Util_FailAfterBlockForTest = True Else Util_FailAfterChartForTest = True
        On Error Resume Next
        Scatter_Insert ws, d, "scatter"
        en = Err.Number
        On Error GoTo 0
        Util_FailAfterBlockForTest = False: Util_FailAfterChartForTest = False
        AssertEqual vbObjectError + k, en
        AssertEqual 0, ws.ChartObjects.Count
        For Each nm In wb.Names
            AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
        Next
    Next
    Teardown wb
End Sub

Public Sub Test_Core_Insert_Scatter_ThroughApi()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:D5").Select
    Core_Insert "scatter"
    Set cht = Core_ActiveChart()
    AssertEqual "scatter", Registry_ChartType(cht)
    AssertTrue Core_Supports("trend") And Core_Supports("quad") And Core_Supports("labels") And Core_Supports("legend")
    AssertFalse Core_Supports("numfmt") Or Core_Supports("gap") Or Core_Supports("totals") Or Core_Supports("samescale")
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), cht.SeriesCollection(1).MarkerBackgroundColor
    AssertEqual RGB(127, 127, 127), cht.SeriesCollection(2).MarkerBackgroundColor
    Core_ClearLastNotice
    Core_SetOption "numfmt", "dec1"
    AssertEqual "'Formato' não se aplica a este tipo de gráfico.", Core_LastNotice()
    Core_SetToggle "trend", True
    AssertTrue Core_GetToggle("trend"): AssertEqual 1, cht.SeriesCollection(1).Trendlines.Count
    Core_SetOption "legend", "none"
    AssertFalse cht.HasLegend
    Core_SetOption "legend", "right"
    AssertEqual 2, cht.Legend.LegendEntries.Count, "religar a legenda nao traz as tendencias de volta"
    Core_SetToggle "labels", False
    AssertFalse cht.SeriesCollection(1).HasDataLabels
    Core_SetToggle "quad", True
    AssertEqual xlAxisCrossesCustom, cht.Axes(xlValue).Crosses
    Teardown wb
End Sub

' Nome vazio sem "0", data no formato da celula, grupos "A" e "a" no MESMO grupo (a comparacao de texto do Excel ignora maiusculas: o ponto
' nao pode entrar em duas series), grupo logico com o texto do Excel e nome longo
' (quebrado pelo Excel em 2 linhas) centrado na altura do ponto.
Public Sub Test_Scatter_NamesAndGroups_BlankDateCaseBooleanLongName()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, k As Long, i As Long, hits As Long, v As Variant, boolText As String, lbl As DataLabel, p As Point
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Produto", "Preço", "Margem", "Grupo")
    ws.Range("A2:D2").Value = Array("P1", 10, 0.2, "A")
    ws.Range("B3:D3").Value = Array(20, 0.35, "a")                          ' nome vazio; grupo "a" = "A"
    ws.Range("A4").Value = DateSerial(2026, 3, 2): ws.Range("A4").NumberFormat = "dd/mm/yyyy"
    ws.Range("B4:D4").Value = Array(30, 0.25, "B")
    ws.Range("A5:C5").Value = Array("Produto de nome bem comprido", 40, 0.5): ws.Range("D5").Value = True
    ws.Range("F1").Formula = "=TRUE&"""""                                   ' o texto que o Excel da ao logico (depende da localidade)
    boolText = CStr(ws.Range("F1").Value): ws.Range("F1").ClearContents
    Set d = ChartData_FromRange(ws.Range("A1:D5"))
    Set cht = Scatter_Insert(ws, d, "scatter").Chart
    AssertEqual 3, cht.SeriesCollection.Count, "A e a sao o mesmo grupo"
    AssertEqual "A", cht.SeriesCollection(1).Name: AssertEqual "B", cht.SeriesCollection(2).Name
    AssertEqual boolText, cht.SeriesCollection(3).Name, "grupo logico com o texto do Excel"
    For i = 1 To 4
        hits = 0
        For k = 1 To cht.SeriesCollection.Count
            v = cht.SeriesCollection(k).Values
            If Not (IsError(v(i)) Or IsEmpty(v(i))) Then hits = hits + 1
        Next
        AssertEqual 1, hits, "ponto " & i & " em exatamente uma serie"
    Next
    AssertEqual "", cht.SeriesCollection(1).Points(2).DataLabel.Text, "nome vazio: rotulo vazio (sem 0)"
    AssertEqual Format$(DateSerial(2026, 3, 2), "dd/mm/yyyy"), cht.SeriesCollection(2).Points(3).DataLabel.Text, "data no formato da celula"
    ForceLayout cht
    Set lbl = cht.SeriesCollection(3).Points(4).DataLabel
    AssertEqual "Produto de nome bem comprido", lbl.Text
    Set p = cht.SeriesCollection(3).Points(4)
    AssertTrue Abs((lbl.Top + lbl.Height / 2) - (p.Top + p.Height / 2)) < 8, "nome longo na altura do ponto"
    AssertTrue lbl.Left >= p.Left + p.Width / 2, "nome longo a direita do ponto"
    CloseTemp wb
End Sub

' Falha no meio de ligar a tendencia (gancho depois da tendencia da 2a serie): nada fica pela metade (nenhuma tendencia, opcao "0"), nenhuma
' entrada de tendencia na legenda e o erro chega ao chamador. Sem o gancho, ligar de novo funciona.
Public Sub Test_Scatter_Trend_FailMidway_RollsBack()
    Dim wb As Workbook, d As clsChartData, cht As Chart, en As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Scatter_Insert(d.Source.Worksheet, d, "scatter").Chart
    Util_FailThinSyncForTest = "scatter-trend:2"
    On Error Resume Next
    Scatter_SetTrend cht, True
    en = Err.Number
    On Error GoTo 0
    Util_FailThinSyncForTest = ""
    AssertEqual vbObjectError + 9996, en, "o erro chega ao chamador"
    AssertEqual 0, cht.SeriesCollection(1).Trendlines.Count, "a tendencia da 1a serie foi desfeita"
    AssertEqual 0, cht.SeriesCollection(2).Trendlines.Count
    AssertEqual "0", Registry_GetOption(cht, "trend")
    AssertEqual 2, cht.Legend.LegendEntries.Count, "sem entrada de tendencia na legenda"
    Scatter_SetTrend cht, True
    AssertEqual 1, cht.SeriesCollection(1).Trendlines.Count: AssertEqual 1, cht.SeriesCollection(2).Trendlines.Count
    AssertEqual "1", Registry_GetOption(cht, "trend"): AssertEqual 2, cht.Legend.LegendEntries.Count
    CloseTemp wb
End Sub

' Calculo manual: o bloco nasce calculado (medias, colunas por grupo) e os quadrantes usam as medias certas.
Public Sub Test_Scatter_ManualCalc_BlockAndQuadrants()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range, prevCalc As Long, en As Long, ed As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    prevCalc = Application.Calculation                  ' (lido com uma pasta aberta: sem pasta o Excel nao devolve o modo)
    Application.Calculation = xlCalculationManual
    On Error GoTo Fail
    Set cht = Scatter_Insert(d.Source.Worksheet, d, "scatter").Chart
    Set blk = Registry_Block(cht)
    AssertAlmost 25, CDbl(blk.Cells(1, 5).Value), 0.000001, "AvgX em manual"
    AssertEqual 20, Block_Body(blk).Cells(3, Scatter_XgCol(blk, 2)).Value, "coluna do grupo B em manual"
    Scatter_SetQuad cht, True
    AssertAlmost 25, CDbl(cht.Axes(xlCategory).CrossesAt), 0.000001
Fail:
    en = Err.Number: ed = Err.Description
    Application.Calculation = prevCalc
    CloseTemp wb
    If en <> 0 Then Err.Raise en, "Test_Scatter_ManualCalc_BlockAndQuadrants", ed
End Sub

' Bolhas: a paleta pinta o preenchimento (MarkerBackgroundColor nao vale para bolhas: erro 1004, medido) e refaz o contraste dos rotulos
' ponto a ponto, sem recriar o rotulo que o usuario apagou.
Public Sub Test_Bubble_Palette_FillAndLabelContrast_KeepsDeletedLabel()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb, True, True)
    Set cht = Scatter_Insert(d.Source.Worksheet, d, "bubble").Chart
    Set s = cht.SeriesCollection(1)
    s.Points(3).HasDataLabel = False                            ' P3 (grupo A): apagado pelo usuario
    Scatter_ApplyColors cht, "cinza"
    AssertEqual RGB(64, 64, 64), s.Format.Fill.ForeColor.RGB
    AssertEqual RGB(127, 127, 127), cht.SeriesCollection(2).Format.Fill.ForeColor.RGB
    AssertEqual Style_ContrastFont(RGB(64, 64, 64)), s.Points(1).DataLabel.Font.Color, "contraste com a cor nova"
    AssertFalse s.Points(3).HasDataLabel, "o rotulo apagado continua apagado"
    Scatter_ApplyColors cht, "padrao"
    AssertEqual Style_ContrastFont(RGB(31, 56, 100)), s.Points(1).DataLabel.Font.Color
    AssertEqual Style_ContrastFont(RGB(68, 114, 196)), cht.SeriesCollection(2).Points(2).DataLabel.Font.Color
    CloseTemp wb
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Revisao 1 da Tarefa 2
' ---------------------------------------------------------------------------------------------------------------------

' As series apontam para o BLOCO (sempre n linhas): a linha inserida no meio da tabela ou apagada so aparece nas referencias do bloco. A
' extensao da tabela (linha da primeira ate a da ultima referencia) muda e "Alterar dados" e pedido; um refresh comum e uma linha inserida
' ACIMA da tabela nao avisam.
Public Sub Test_Scatter_SourceMismatch_InsertDeleteMiddleRow()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Scatter_Insert(ws, d, "scatter").Chart
    AssertEqual 4, Scatter_SourceRowsSpanned(cht)
    AssertFalse Core_ChartSourceMismatch(cht), "recem-inserido"
    ws.Range("B2").Value = 15: cht.Refresh: Core_SafeRefresh cht
    AssertFalse Core_ChartSourceMismatch(cht), "refresh comum nao avisa"
    ws.Rows(1).Insert                                             ' acima da tabela: as duas pontas descem juntas
    AssertEqual 4, Scatter_SourceRowsSpanned(cht)
    AssertFalse Core_ChartSourceMismatch(cht), "linha acima da tabela nao avisa"
    ws.Rows(4).Insert                                             ' no meio (entre P1 e P2)
    AssertEqual 5, Scatter_SourceRowsSpanned(cht)
    AssertTrue Core_ChartSourceMismatch(cht), "linha inserida no meio"
    Core_ClearLastNotice
    Core_SafeRefresh cht
    AssertTrue InStr(Core_LastNotice(), "mudou de tamanho") > 0, "aviso: " & Core_LastNotice()
    ws.Rows(4).Delete
    AssertFalse Core_ChartSourceMismatch(cht), "de volta ao tamanho original"
    ws.Rows(4).Delete                                             ' apaga a linha de P2 (meio)
    AssertEqual 3, Scatter_SourceRowsSpanned(cht)
    AssertTrue Core_ChartSourceMismatch(cht), "linha do meio apagada"
    ws.Rows(3).Delete                                             ' apaga a linha de P1 (a primeira): a referencia vira #REF!
    AssertEqual -1, Scatter_SourceRowsSpanned(cht)
    AssertTrue Core_ChartSourceMismatch(cht), "primeira linha apagada"
    Teardown wb
End Sub

' Bolhas: o nome fica no centro com a fonte de contraste so quando cabe no diametro da bolha (largura estimada do texto); senao vai para a
' direita em cinza-escuro (como na dispersao). Reavaliado a cada refresh; o rotulo apagado pelo usuario continua apagado.
Public Sub Test_Bubble_LabelFitsDiameter_CenterOrRightDarkGray()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, sa As Series, sb As Series, p As Point
    Set wb = NewTempWorkbook(): Set d = Fixture(wb, True, True): Set ws = d.Source.Worksheet
    ws.Range("A3").Value = "Produto de nome bem comprido"                  ' P2: a maior bolha (120), grupo B (escuro: fonte branca se coubesse)
    Set cht = Scatter_Insert(ws, d, "bubble").Chart
    Set sa = cht.SeriesCollection(1): Set sb = cht.SeriesCollection(2)
    Set p = sa.Points(1)                                                     ' P1 (50): "P1" cabe
    AssertTrue Labels_EstimateWidth("P1") <= p.Width, "premissa: P1 cabe (" & p.Width & ")"
    AssertEqual xlLabelPositionCenter, p.DataLabel.Position
    AssertEqual Style_ContrastFont(CLng(sa.Format.Fill.ForeColor.RGB)), p.DataLabel.Font.Color, "contraste dentro da bolha"
    Set p = sb.Points(2)                                                     ' nome longo na maior bolha: nao cabe
    AssertTrue Labels_EstimateWidth(p.DataLabel.Text) > p.Width, "premissa: o nome longo nao cabe (" & p.Width & ")"
    AssertEqual xlLabelPositionRight, p.DataLabel.Position, "nome que nao cabe vai para a direita"
    AssertEqual RGB(64, 64, 64), p.DataLabel.Font.Color, "cinza-escuro fora da bolha (nunca branco sobre branco)"
    Set p = sb.Points(4)                                                     ' P4 (20): bolha pequena, "P4" nao cabe
    AssertTrue Labels_EstimateWidth("P4") > p.Width, "premissa: P4 nao cabe (" & p.Width & ")"
    AssertEqual xlLabelPositionRight, p.DataLabel.Position
    AssertEqual RGB(64, 64, 64), p.DataLabel.Font.Color
    sa.Points(3).HasDataLabel = False                                        ' P3 apagado pelo usuario
    ws.Range("D5").Value = 150                                               ' P4 cresce: passa a caber
    cht.Refresh: Core_SafeRefresh cht
    Set p = sb.Points(4)
    AssertTrue Labels_EstimateWidth("P4") <= p.Width, "premissa: P4 agora cabe (" & p.Width & ")"
    AssertEqual xlLabelPositionCenter, p.DataLabel.Position, "reavaliado no refresh"
    AssertEqual Style_ContrastFont(CLng(sb.Format.Fill.ForeColor.RGB)), p.DataLabel.Font.Color
    AssertFalse sa.Points(3).HasDataLabel, "o rotulo apagado continua apagado"
    Scatter_ApplyColors cht, "cinza"                                         ' a paleta respeita a posicao: fora continua cinza-escuro
    AssertEqual RGB(64, 64, 64), sb.Points(2).DataLabel.Font.Color
    AssertEqual Style_ContrastFont(RGB(127, 127, 127)), sb.Points(4).DataLabel.Font.Color
    AssertFalse sa.Points(3).HasDataLabel
    Teardown wb
End Sub

' Celula de grupo com erro: o bloco cai em "(sem grupo)" (IFERROR), como GroupNames supoe; o ponto e plotado nessa serie.
Public Sub Test_Scatter_ErrorGroupCell_GoesToNoGroup()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, v As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("D3").Formula = "=1/0"                                          ' P2: grupo com erro
    Set cht = Scatter_Insert(ws, d, "scatter").Chart
    AssertEqual 3, cht.SeriesCollection.Count
    AssertEqual "(sem grupo)", cht.SeriesCollection(2).Name
    v = cht.SeriesCollection(2).Values
    AssertTrue Not (IsError(v(2)) Or IsEmpty(v(2))), "P2 plotado em (sem grupo)"
    AssertAlmost 0.35, CDbl(v(2)), 0.000001
    CloseTemp wb
End Sub

' Falha no meio de ligar os quadrantes (gancho no segundo eixo): o primeiro eixo volta ao automatico e a opcao nao e gravada; o erro chega ao
' chamador. Sem o gancho, ligar de novo funciona.
Public Sub Test_Scatter_Quad_FailMidway_RollsBack()
    Dim wb As Workbook, d As clsChartData, cht As Chart, en As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Scatter_Insert(d.Source.Worksheet, d, "scatter").Chart
    Util_FailThinSyncForTest = "scatter-quad:2"
    On Error Resume Next
    Scatter_SetQuad cht, True
    en = Err.Number
    On Error GoTo 0
    Util_FailThinSyncForTest = ""
    AssertEqual vbObjectError + 9996, en, "o erro chega ao chamador"
    AssertEqual xlAxisCrossesAutomatic, cht.Axes(xlCategory).Crosses, "o primeiro eixo foi desfeito"
    AssertEqual msoLineSolid, cht.Axes(xlCategory).Format.Line.DashStyle
    AssertEqual xlAxisCrossesAutomatic, cht.Axes(xlValue).Crosses
    AssertEqual "0", Registry_GetOption(cht, "quad")
    Scatter_SetQuad cht, True
    AssertEqual xlAxisCrossesCustom, cht.Axes(xlCategory).Crosses: AssertEqual xlAxisCrossesCustom, cht.Axes(xlValue).Crosses
    AssertEqual "1", Registry_GetOption(cht, "quad")
    CloseTemp wb
End Sub

' Nomes de grupo que o Excel recusaria ou confundiria: "=x" (Series.Name o leria como formula: erro 1004, medido), "tqq_g" (seria tomada por
' serie auxiliar) e um nome de 300 caracteres (literal de formula e Series.Name: no maximo 255, medido). O grafico nasce, cada ponto fica na sua
' serie e a legenda (com tendencia) mostra os 4 grupos.
Public Sub Test_Scatter_SpecialGroupNames_EqualsHelperPrefixLong()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, i As Long, k As Long, hits As Long, v As Variant, longName As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    longName = String(250, "a") & String(50, "b")
    ws.Range("D2").Value = "'=x": ws.Range("D3").Value = "tqq_g": ws.Range("D4").Value = longName: ws.Range("D5").Value = "B"
    Set cht = Scatter_Insert(ws, d, "scatter").Chart
    AssertEqual 4, cht.SeriesCollection.Count: AssertEqual 4, Scatter_SeriesCount(cht), "nenhum grupo tomado por auxiliar"
    AssertEqual " =x", cht.SeriesCollection(1).Name, "prefixo de espaco: o Excel nao le como formula"
    AssertEqual " tqq_g", cht.SeriesCollection(2).Name
    AssertEqual Left$(longName, 255), cht.SeriesCollection(3).Name, "nome limitado a 255 caracteres"
    For i = 1 To 4
        hits = 0
        For k = 1 To 4
            v = cht.SeriesCollection(k).Values
            If Not (IsError(v(i)) Or IsEmpty(v(i))) Then hits = hits + 1
        Next
        AssertEqual 1, hits, "ponto " & i & " em exatamente uma serie"
    Next
    Scatter_SetTrend cht, True
    AssertEqual 4, cht.Legend.LegendEntries.Count, "4 grupos na legenda, sem as tendencias"
    AssertEqual 1, cht.SeriesCollection(2).Trendlines.Count, "o grupo tqq_g tambem ganha tendencia"
    CloseTemp wb
End Sub

' As medias dos quadrantes contam so os pontos PLOTADOS: Y invalido (dispersao) e tamanho <= 0 (bolhas) ficam fora da media de X e de Y.
Public Sub Test_Scatter_QuadAverages_OnlyPlottedPoints()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Worksheet.Range("C3").Value = "x"                              ' P2: Y invalido
    Set cht = Scatter_Insert(d.Source.Worksheet, d, "scatter").Chart
    Set blk = Registry_Block(cht)
    AssertAlmost (10 + 30 + 40) / 3, CDbl(blk.Cells(1, 5).Value), 0.000001, "AvgX sem P2"
    AssertAlmost (0.2 + 0.25 + 0.5) / 3, CDbl(blk.Cells(1, 6).Value), 0.000001, "AvgY sem P2"
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb, True, True)
    d.Source.Worksheet.Range("D5").Value = -20                              ' P4: tamanho negativo
    Set cht = Scatter_Insert(d.Source.Worksheet, d, "bubble").Chart
    Set blk = Registry_Block(cht)
    AssertAlmost (10 + 20 + 30) / 3, CDbl(blk.Cells(1, 5).Value), 0.000001, "AvgX sem a bolha negativa"
    AssertAlmost (0.2 + 0.35 + 0.25) / 3, CDbl(blk.Cells(1, 6).Value), 0.000001, "AvgY sem a bolha negativa"
    CloseTemp wb
End Sub

' Onda A, item B: os numeros dos eixos seguem o formato das celulas de origem de X e Y (na demonstracao a margem em % aparecia 0,3 / 0,325),
' na insercao e no refresh; formato Geral mantem o padrao. Dispersao e bolhas.
Public Sub Test_Scatter_AxisNumberFormat_FollowsSource()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, k As Variant, d As clsChartData, axX As Axis, axY As Axis
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Loja", "Margem", "Ticket", "Vendas")
    ws.Range("A2:D2").Value = Array("L1", 0.3, 33.5, 10): ws.Range("A3:D3").Value = Array("L2", 0.31, 32.5, 20)
    ws.Range("A4:D4").Value = Array("L3", 0.325, 33, 30): ws.Range("A5:D5").Value = Array("L4", 0.33, 34.25, 40)
    ws.Range("B2:B5").NumberFormat = "0%"
    ws.Range("C2:C5").NumberFormat = "0.0"
    For Each k In Array("scatter", "bubble")
        Set d = ChartData_FromRange(ws.Range(IIf(k = "scatter", "A1:C5", "A1:D5")))
        Set cht = Scatter_Insert(ws, d, CStr(k)).Chart
        Set axX = cht.Axes(xlCategory, xlPrimary): Set axY = cht.Axes(xlValue, xlPrimary)
        AssertEqual ws.Range("B2").NumberFormatLocal, axX.TickLabels.NumberFormatLocal, k & ": eixo X no formato da coluna X"
        AssertEqual ws.Range("C2").NumberFormatLocal, axY.TickLabels.NumberFormatLocal, k & ": eixo Y no formato da coluna Y"
        ws.Range("C2:C5").NumberFormat = "0.00"
        Core_SafeRefresh cht
        AssertEqual ws.Range("C2").NumberFormatLocal, axY.TickLabels.NumberFormatLocal, k & ": o refresh acompanha o formato novo"
        ws.Range("C2:C5").NumberFormat = "General"
        Core_SafeRefresh cht
        AssertTrue axY.TickLabels.NumberFormatLinked, k & ": Geral volta ao padrao (ligado a origem)"
        ws.Range("C2:C5").NumberFormat = "0.0"
        ws.ChartObjects(1).Delete
        CalcSheet_Cleanup wb
    Next
    Teardown wb
End Sub

' Onda A, item C: faixa "bonita" (pura): dados longe do zero e faixa estreita ganham minimo/maximo em volta dos dados (5% de folga, passo
' 1/2/2,5/5 x 10^k, no maximo 8 intervalos, passo >= a resolucao do formato); zero dentro da faixa ou faixa larga = escala do Excel.
Public Sub Test_Scatter_NiceRange_Pure()
    Dim mn As Double, mx As Double, st As Double
    AssertTrue Scatter_NiceRange(20, 50, 0, 0.05, mn, mx, st), "20..50 aperta"
    AssertAlmost 15, mn, 0.000001: AssertAlmost 55, mx, 0.000001: AssertAlmost 5, st, 0.000001
    AssertTrue Scatter_NiceRange(0.3, 0.33, 0, 0.05, mn, mx, st), "margens de 30% a 33%"
    AssertAlmost 0.295, mn, 0.0000001: AssertAlmost 0.335, mx, 0.0000001: AssertAlmost 0.005, st, 0.0000001
    AssertTrue Scatter_NiceRange(0.3, 0.33, 0.01, 0.05, mn, mx, st), "formato 0%: passo de pelo menos 1 ponto percentual"
    AssertAlmost 0.29, mn, 0.0000001: AssertAlmost 0.34, mx, 0.0000001: AssertAlmost 0.01, st, 0.0000001
    AssertFalse Scatter_NiceRange(-5, 50, 0, 0.05, mn, mx, st), "zero dentro da faixa: escala do Excel"
    AssertFalse Scatter_NiceRange(10, 50, 0, 0.05, mn, mx, st), "faixa larga (10..50): escala do Excel"
    AssertTrue Scatter_NiceRange(-50, -20, 0, 0.05, mn, mx, st), "negativos: espelho"
    AssertAlmost -55, mn, 0.000001: AssertAlmost -15, mx, 0.000001
    AssertTrue Scatter_NiceRange(30, 30, 0, 0.05, mn, mx, st), "um valor so"
    AssertTrue mn < 30 And mx > 30 And mn > 0, "um valor so: em volta dele (" & mn & ".." & mx & ")"
    AssertTrue Scatter_NiceRange(1, 1.4, 0, 0.05, mn, mx, st)
    AssertTrue mn >= 0, "nunca cruza o zero pelo arredondamento"
End Sub

' Onda A, item C: pontos entre 20 e 50 nao ficam amontoados num canto (o Excel comecaria os eixos em zero): a escala aperta em volta deles,
' e gravada (xscale/yscale), os quadrantes continuam nas medias, os dados que passam a incluir o zero (ou a ficar largos) devolvem a escala
' do Excel, e um eixo que o usuario fixou a mao nao e sobrescrito. Bolhas: a maior bolha fica inteira dentro da area de plotagem.
Public Sub Test_Scatter_TightAxes_Between20And50()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, axX As Axis, axY As Axis, p As Point, d As clsChartData, s As Series, i As Long, k As Variant
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Loja", "Ticket", "Margem", "Vendas")
    ws.Range("A2:D2").Value = Array("L1", 20, 21, 50): ws.Range("A3:D3").Value = Array("L2", 28, 30, 120)
    ws.Range("A4:D4").Value = Array("L3", 35, 38, 80): ws.Range("A5:D5").Value = Array("L4", 42, 45, 20)
    ws.Range("A6:D6").Value = Array("L5", 50, 49, 60)
    Set d = ChartData_FromRange(ws.Range("A1:C6"))
    Set cht = Scatter_Insert(ws, d, "scatter").Chart
    Set axX = cht.Axes(xlCategory, xlPrimary): Set axY = cht.Axes(xlValue, xlPrimary)
    For Each k In Array(1, 2)
        With cht.Axes(IIf(k = 1, xlCategory, xlValue), xlPrimary)
            AssertFalse .MinimumScaleIsAuto Or .MaximumScaleIsAuto, "eixo " & k & ": escala do ThinQQ"
            AssertTrue .MinimumScale > 0 And .MinimumScale <= 20 And .MaximumScale >= 50, "eixo " & k & ": cobre os dados sem o zero (" & .MinimumScale & ".." & .MaximumScale & ")"
            AssertTrue 29 / (.MaximumScale - .MinimumScale) >= 0.7, "eixo " & k & ": os pontos ocupam a maior parte do eixo (" & .MinimumScale & ".." & .MaximumScale & ")"
        End With
    Next
    AssertTrue Len(Registry_GetOption(cht, "xscale", "")) > 0 And Len(Registry_GetOption(cht, "yscale", "")) > 0, "escala gravada"
    ForceLayout cht
    Set p = cht.SeriesCollection(1).Points(1)                   ' L1 (20; 21): regra de tres no eixo real
    AssertAlmost XOf(cht, 20), p.Left + p.Width / 2, 2, "centro do marcador em x(20)"
    AssertAlmost YOf(cht, 21), p.Top + p.Height / 2, 2, "centro do marcador em y(21)"
    AssertTrue p.Left + p.Width / 2 > cht.PlotArea.InsideLeft + 2, "L1 nao encostado no canto"
    Scatter_SetQuad cht, True
    AssertAlmost 35, CDbl(axX.CrossesAt), 0.000001, "quadrante na media de X"
    AssertTrue axX.CrossesAt > axX.MinimumScale And axX.CrossesAt < axX.MaximumScale, "a linha do quadrante fica dentro da escala"
    AssertFalse axX.MinimumScaleIsAuto, "os quadrantes mantem a escala"
    ws.Range("B6").Value = 80                                   ' X de 20 a 80: faixa larga, escala do Excel
    Core_SafeRefresh cht
    AssertTrue axX.MinimumScaleIsAuto And axX.MaximumScaleIsAuto, "faixa larga: escala do Excel"
    AssertEqual "", Registry_GetOption(cht, "xscale", ""), "opcao limpa"
    AssertAlmost 41, CDbl(axX.CrossesAt), 0.000001, "quadrante acompanha"
    ws.Range("B6").Value = 50
    Core_SafeRefresh cht
    AssertFalse axX.MinimumScaleIsAuto, "de volta ao aperto"
    ws.Range("B2").Value = -10                                  ' zero dentro da faixa
    Core_SafeRefresh cht
    AssertTrue axX.MinimumScaleIsAuto And axX.MaximumScaleIsAuto, "zero dentro da faixa: escala do Excel"
    ws.Range("B2").Value = 20
    Core_SafeRefresh cht
    axY.MinimumScale = 0                                        ' o usuario fixa o eixo Y
    Core_SafeRefresh cht
    AssertEqual 0, axY.MinimumScale, "eixo do usuario fica"
    ws.Range("C2").Value = 25
    Core_SafeRefresh cht
    AssertEqual 0, axY.MinimumScale, "eixo do usuario fica depois de mudar os dados"
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    Set d = ChartData_FromRange(ws.Range("A1:D6"))
    Set cht = Scatter_Insert(ws, d, "bubble").Chart
    AssertFalse cht.Axes(xlCategory, xlPrimary).MinimumScaleIsAuto, "bolhas: escala do ThinQQ"
    ForceLayout cht
    For Each s In cht.SeriesCollection
        For i = 1 To s.Points.Count
            Set p = s.Points(i)
            If p.Width > 0 Then
                AssertTrue p.Left >= cht.PlotArea.InsideLeft - 1 And p.Left + p.Width <= cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth + 1, "bolha " & i & " inteira na horizontal (" & p.Left & "+" & p.Width & ")"
                AssertTrue p.Top >= cht.PlotArea.InsideTop - 1 And p.Top + p.Height <= cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight + 1, "bolha " & i & " inteira na vertical (" & p.Top & "+" & p.Height & ")"
            End If
        Next
    Next
    Teardown wb
End Sub

' Onda A, item C: a escala do eixo e transacional (uma falha depois de mexer no eixo devolve o automatico e a opcao) e nunca derruba a
' insercao; sem a falha, o refresh seguinte aperta.
Public Sub Test_Scatter_TightAxes_FailMidway_RollsBack()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, d As clsChartData, en As Long, es As String, ed As String
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Loja", "Ticket", "Margem")
    ws.Range("A2:C2").Value = Array("L1", 20, 21): ws.Range("A3:C3").Value = Array("L2", 50, 49)
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    Util_FailThinSyncForTest = "scatter-scale:1"
    Set cht = Scatter_Insert(ws, d, "scatter").Chart
    AssertTrue cht.Axes(xlCategory, xlPrimary).MinimumScaleIsAuto And cht.Axes(xlCategory, xlPrimary).MaximumScaleIsAuto, "falha no eixo X: automatico de volta"
    AssertTrue cht.Axes(xlCategory, xlPrimary).MajorUnitIsAuto, "passo automatico de volta"
    AssertEqual "", Registry_GetOption(cht, "xscale", ""), "opcao como estava"
    AssertFalse cht.Axes(xlValue, xlPrimary).MinimumScaleIsAuto, "o eixo Y (sem falha) aperta"
    Util_FailThinSyncForTest = ""
    Core_SafeRefresh cht
    AssertFalse cht.Axes(xlCategory, xlPrimary).MinimumScaleIsAuto, "sem a falha o refresh aperta"
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailThinSyncForTest = ""
    Teardown wb
    If en <> 0 Then Err.Raise en, es, ed
End Sub

' Onda A, rodada 1 (Importante 2): o formato que o USUARIO poe no eixo nao e revertido pelo refresh. O ThinQQ grava o formato que aplicou
' (xfmt/yfmt) e so reescreve enquanto o eixo esta ligado a origem ou ainda mostra o formato gravado; devolvido a "ligado a origem", o ThinQQ
' volta a seguir a coluna. Uma falha no meio (gancho) devolve o eixo e a opcao.
Public Sub Test_Scatter_AxisNumberFormat_UserFormatKept()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, d As clsChartData, axX As Axis, userFmt As String, en As Long, es As String, ed As String, keep As String
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Loja", "Margem", "Ticket")
    ws.Range("A2:C2").Value = Array("L1", 0.3, 10): ws.Range("A3:C3").Value = Array("L2", 0.33, 30)
    ws.Range("B2:B3").NumberFormat = "0%"
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    Set cht = Scatter_Insert(ws, d, "scatter").Chart
    Set axX = cht.Axes(xlCategory, xlPrimary)
    AssertEqual ws.Range("B2").NumberFormatLocal, axX.TickLabels.NumberFormatLocal, "formato da origem"
    userFmt = "0" & Application.International(xlDecimalSeparator) & "0%"
    axX.TickLabels.NumberFormatLocal = userFmt                  ' o usuario muda o formato do eixo
    Core_SafeRefresh cht
    Core_SafeRefresh cht
    AssertEqual userFmt, axX.TickLabels.NumberFormatLocal, "o formato do usuario fica"
    ws.Range("B2:B3").NumberFormat = "0.00%"
    Core_SafeRefresh cht
    AssertEqual userFmt, axX.TickLabels.NumberFormatLocal, "mesmo com a origem mudando"
    axX.TickLabels.NumberFormatLinked = True                    ' o usuario devolve a "ligado a origem"
    Core_SafeRefresh cht
    AssertEqual ws.Range("B2").NumberFormatLocal, axX.TickLabels.NumberFormatLocal, "o ThinQQ volta a seguir a coluna"
    keep = Registry_GetOption(cht, "xfmt", "")
    Core_SafeRefresh cht
    AssertEqual keep, Registry_GetOption(cht, "xfmt", ""), "dois refresh: nada muda"
    AssertEqual ws.Range("B2").NumberFormatLocal, axX.TickLabels.NumberFormatLocal
    ws.Range("B2:B3").NumberFormat = "0%"
    Util_FailThinSyncForTest = "scatter-fmt:1"
    Core_SafeRefresh cht
    Util_FailThinSyncForTest = ""
    AssertEqual Replace(ws.Range("B2").NumberFormatLocal, "0%", "0" & Application.International(xlDecimalSeparator) & "00%"), axX.TickLabels.NumberFormatLocal, "falha: o eixo volta ao formato anterior"
    AssertEqual keep, Registry_GetOption(cht, "xfmt", ""), "falha: a opcao fica"
    Core_SafeRefresh cht
    AssertEqual ws.Range("B2").NumberFormatLocal, axX.TickLabels.NumberFormatLocal, "sem a falha, segue a coluna"
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailThinSyncForTest = ""
    Teardown wb
    If en <> 0 Then Err.Raise en, es, ed
End Sub

' Onda A, rodada 1 (menor d): "[Red]0%" na origem: o passo da escala apertada respeita a resolucao do formato (1 ponto percentual); antes o
' "d" de [Red] fazia o formato parecer data (resolucao 0, passo de 0,5% que repetiria "30%").
Public Sub Test_Scatter_TightAxes_BracketFormatUnit()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Loja", "Margem", "Ticket")
    ws.Range("A2:C2").Value = Array("L1", 0.3, 20): ws.Range("A3:C3").Value = Array("L2", 0.33, 50)
    ws.Range("B2:B3").NumberFormat = "[Red]0%"
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    Set cht = Scatter_Insert(ws, d, "scatter").Chart
    AssertAlmost 0.01, cht.Axes(xlCategory, xlPrimary).MajorUnit, 0.0000001, "passo de 1 ponto percentual"
    Teardown wb
End Sub

' Onda A, rodada 2 (menor 1): o eixo nem sempre devolve o formato como foi gravado (medido: #.##0,00 "R$" volta como #.##0,00 R$; [$R$-416]
' volta como [$R$-pt-BR]). O ThinQQ grava o que o EIXO devolve e continua seguindo a coluna de origem quando ela muda.
Public Sub Test_Scatter_AxisNumberFormat_ReadBackNormalized_KeepsFollowing()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, d As clsChartData, axX As Axis
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Loja", "Ticket", "Margem")
    ws.Range("A2:C2").Value = Array("L1", 300, 0.3): ws.Range("A3:C3").Value = Array("L2", 330, 0.33)
    ws.Range("B2:B3").NumberFormat = "#,##0.00 ""R$"""
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    Set cht = Scatter_Insert(ws, d, "scatter").Chart
    Set axX = cht.Axes(xlCategory, xlPrimary)
    AssertFalse axX.TickLabels.NumberFormatLinked, "formato da origem aplicado"
    Core_SafeRefresh cht
    ws.Range("B2:B3").NumberFormat = "0.0"
    Core_SafeRefresh cht
    AssertEqual ws.Range("B2").NumberFormatLocal, axX.TickLabels.NumberFormatLocal, "continua seguindo a coluna"
    ws.Range("B2:B3").NumberFormat = "[$R$-416] #,##0.00"
    Core_SafeRefresh cht
    ws.Range("B2:B3").NumberFormat = "0.00"
    Core_SafeRefresh cht
    AssertEqual ws.Range("B2").NumberFormatLocal, axX.TickLabels.NumberFormatLocal, "continua seguindo (moeda com codigo de localidade)"
    Teardown wb
End Sub

' Onda B (B2.6): Scatter_SyncBubbleLabels escreve com os eventos desligados e barra a reentrada (cada escrita disparava Calculate e o refresh
' entrava de novo: convergia, mas fundo). Uma chamada = uma passada que faz o trabalho; EnableEvents volta ao que era (ligado ou desligado).
Public Sub Test_Bubble_SyncLabels_NoReentry_EventsRestored()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, i As Long, prevEv As Boolean, en As Long, es As String, ed As String
    prevEv = Application.EnableEvents
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set d = Fixture(wb, True, True)
    d.Source.Select
    Core_Insert "bubble"
    Set cht = Core_ActiveChart()
    Application.EnableEvents = False                                         ' prepara sem refresh: todos os rotulos fora do lugar
    For Each s In cht.SeriesCollection
        For i = 1 To s.Points.Count
            If s.Points(i).HasDataLabel Then s.Points(i).DataLabel.Position = xlLabelPositionLeft
        Next
    Next
    Application.EnableEvents = True
    Scatter_BubbleSyncRunsForTest True
    Scatter_SyncBubbleLabels cht
    AssertEqual 1, Scatter_BubbleSyncRunsForTest(True), "uma passada (as escritas nao reentram pelo Calculate)"
    AssertTrue Application.EnableEvents, "eventos ligados de volta"
    AssertEqual xlLabelPositionCenter, cht.SeriesCollection(1).Points(1).DataLabel.Position, "a passada fez o trabalho (P1 cabe)"
    AssertEqual xlLabelPositionRight, cht.SeriesCollection(2).Points(4).DataLabel.Position, "P4 (bolha pequena) nao cabe"
    Application.EnableEvents = False
    Scatter_SyncBubbleLabels cht
    AssertFalse Application.EnableEvents, "desligado antes: continua desligado"
    Err.Clear                                                   ' um erro ja tratado (On Error Resume Next) no core nao e falha do teste
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    Application.EnableEvents = prevEv                          ' os eventos voltam em toda saida
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, es, ed
End Sub

' Escala gravada (xscale/yscale) do eixo k igual a do eixo real (minimo|maximo|passo)?
Private Function ScaleMatchesAxis(cht As Chart, k As Long) As Boolean
    Dim p As Variant, ax As Axis
    Set ax = cht.Axes(IIf(k = 1, xlCategory, xlValue), xlPrimary)
    p = Split(Registry_GetOption(cht, IIf(k = 1, "xscale", "yscale"), ""), "|")
    If UBound(p) < 2 Then Exit Function
    If ax.MinimumScaleIsAuto Or ax.MaximumScaleIsAuto Or ax.MajorUnitIsAuto Then Exit Function
    ScaleMatchesAxis = Abs(ax.MinimumScale - Val(p(0))) < 0.000001 And Abs(ax.MaximumScale - Val(p(1))) < 0.000001 And Abs(ax.MajorUnit - Val(p(2))) < 0.000001
End Function

' Revisao final: "Alterar dados" numa dispersao guarda a escala e o formato que o grafico NOVO calculou (xscale/yscale/xfmt/yfmt/xsrc/ysrc), nao
' os do grafico antigo: senao o refresh nao reconhece a propria escala, congela o eixo e o ponto alem do maximo antigo some.
Public Sub Test_Scatter_Rebuild_KeepsNewScaleAndReTightens()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, nc As Chart, axX As Axis, i As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Loja", "Ticket", "Margem")
    ws.Range("E1:G1").Value = Array("Loja", "Ticket", "Margem")
    For i = 0 To 4
        ws.Cells(2 + i, 1).Value = "L" & (i + 1): ws.Cells(2 + i, 2).Value = 20 + 7.5 * i: ws.Cells(2 + i, 3).Value = 21 + 7 * i
        ws.Cells(2 + i, 5).Value = "M" & (i + 1): ws.Cells(2 + i, 6).Value = 100 + 10 * i: ws.Cells(2 + i, 7).Value = 21 + 7 * i
    Next
    ws.Range("F2:F6").NumberFormat = "0.0"
    ws.Range("A1:C6").Select
    Core_Insert "scatter"
    Set cht = Core_ActiveChart()
    AssertTrue ScaleMatchesAxis(cht, 1), "antes: escala X gravada = eixo"
    AssertTrue Core_RebuildFromRange(cht, ws.Range("E1:G6")), "Alterar dados (" & Core_LastNotice() & ")"
    Set nc = Core_ActiveChart(): Set axX = nc.Axes(xlCategory, xlPrimary)
    AssertTrue axX.MinimumScale > 50 And axX.MaximumScale >= 140, "eixo X do grafico novo apertado nos dados novos (" & axX.MinimumScale & ".." & axX.MaximumScale & ")"
    AssertTrue ScaleMatchesAxis(nc, 1), "escala X gravada = eixo do grafico novo (" & Registry_GetOption(nc, "xscale", "") & ")"
    AssertTrue ScaleMatchesAxis(nc, 2), "escala Y gravada = eixo do grafico novo (" & Registry_GetOption(nc, "yscale", "") & ")"
    AssertTrue Len(Registry_GetOption(nc, "xfmt", "")) > 0, "formato X da origem nova gravado"
    ws.Range("F6").Value = 175                                  ' alem do maximo: o ThinQQ aperta de novo
    Core_SafeRefresh nc
    AssertTrue axX.MaximumScale >= 175, "o ponto novo fica dentro do eixo (" & axX.MaximumScale & ")"
    AssertTrue ScaleMatchesAxis(nc, 1), "escala regravada"
    ws.Range("F2:F6").NumberFormat = "0.00"
    Core_SafeRefresh nc
    AssertEqual ws.Range("F2").NumberFormatLocal, axX.TickLabels.NumberFormatLocal, "o formato do eixo continua seguindo a coluna"
    Teardown wb
End Sub

' Revisao final (item 7): a escala e o formato que o usuario fixou a mao num eixo da dispersao continuam depois de "Alterar dados" (como
' vmin/vmax nos outros tipos) e o refresh nao os desfaz.
Public Sub Test_Scatter_Rebuild_KeepsUserFixedScaleAndFormat()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, nc As Chart, i As Long, userFmt As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Loja", "Ticket", "Margem")
    ws.Range("E1:G1").Value = Array("Loja", "Ticket", "Margem")
    For i = 0 To 4
        ws.Cells(2 + i, 1).Value = "L" & (i + 1): ws.Cells(2 + i, 2).Value = 20 + 7.5 * i: ws.Cells(2 + i, 3).Value = 21 + 7 * i
        ws.Cells(2 + i, 5).Value = "M" & (i + 1): ws.Cells(2 + i, 6).Value = 100 + 10 * i: ws.Cells(2 + i, 7).Value = 31 + 5 * i
    Next
    ws.Range("A1:C6").Select
    Core_Insert "scatter"
    Set cht = Core_ActiveChart()
    cht.Axes(xlValue, xlPrimary).MinimumScale = 0               ' o usuario fixa o eixo Y
    cht.Axes(xlValue, xlPrimary).MaximumScale = 100
    userFmt = "0" & Application.International(xlDecimalSeparator) & "00"
    cht.Axes(xlCategory, xlPrimary).TickLabels.NumberFormatLocal = userFmt      ' e o formato do eixo X
    AssertTrue Core_RebuildFromRange(cht, ws.Range("E1:G6")), "Alterar dados (" & Core_LastNotice() & ")"
    Set nc = Core_ActiveChart()
    AssertEqual 0, nc.Axes(xlValue, xlPrimary).MinimumScale, "minimo do usuario mantido"
    AssertEqual 100, nc.Axes(xlValue, xlPrimary).MaximumScale, "maximo do usuario mantido"
    AssertEqual "", Registry_GetOption(nc, "yscale", ""), "eixo Y do usuario: sem escala do ThinQQ gravada"
    AssertEqual userFmt, nc.Axes(xlCategory, xlPrimary).TickLabels.NumberFormatLocal, "formato do usuario mantido"
    AssertTrue ScaleMatchesAxis(nc, 1), "eixo X (escala do ThinQQ) apertado e gravado"
    ws.Range("G6").Value = 60
    Core_SafeRefresh nc
    AssertEqual 0, nc.Axes(xlValue, xlPrimary).MinimumScale, "refresh nao desfaz o minimo do usuario"
    AssertEqual 100, nc.Axes(xlValue, xlPrimary).MaximumScale, "nem o maximo"
    AssertEqual userFmt, nc.Axes(xlCategory, xlPrimary).TickLabels.NumberFormatLocal, "nem o formato"
    Teardown wb
End Sub

' Revisao final (item 7): uma falha ao passar o eixo do usuario para o grafico novo nao derruba "Alterar dados": o eixo novo volta a escala
' que o ThinQQ calculou, com a opcao gravada coerente com ele (o refresh continua apertando).
Public Sub Test_Scatter_Rebuild_KeepUserAxis_FailMidway_RollsBack()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, nc As Chart, i As Long, en As Long, es As String, ed As String
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Loja", "Ticket", "Margem")
    ws.Range("E1:G1").Value = Array("Loja", "Ticket", "Margem")
    For i = 0 To 4
        ws.Cells(2 + i, 1).Value = "L" & (i + 1): ws.Cells(2 + i, 2).Value = 20 + 7.5 * i: ws.Cells(2 + i, 3).Value = 21 + 7 * i
        ws.Cells(2 + i, 5).Value = "M" & (i + 1): ws.Cells(2 + i, 6).Value = 100 + 10 * i: ws.Cells(2 + i, 7).Value = 31 + 5 * i
    Next
    ws.Range("A1:C6").Select
    Core_Insert "scatter"
    Set cht = Core_ActiveChart()
    cht.Axes(xlValue, xlPrimary).MinimumScale = 0
    cht.Axes(xlValue, xlPrimary).MaximumScale = 100
    Util_FailThinSyncForTest = "scatter-keep:2"
    AssertTrue Core_RebuildFromRange(cht, ws.Range("E1:G6")), "Alterar dados nao falha (" & Core_LastNotice() & ")"
    Util_FailThinSyncForTest = ""
    Set nc = Core_ActiveChart()
    AssertTrue ScaleMatchesAxis(nc, 2), "eixo Y volta a escala do ThinQQ, gravada (" & Registry_GetOption(nc, "yscale", "") & ")"
    AssertTrue nc.Axes(xlValue, xlPrimary).MinimumScale > 0, "sem o minimo do usuario pela metade"
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailThinSyncForTest = ""
    Teardown wb
    If en <> 0 Then Err.Raise en, es, ed
End Sub
