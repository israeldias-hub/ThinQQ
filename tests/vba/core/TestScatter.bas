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
