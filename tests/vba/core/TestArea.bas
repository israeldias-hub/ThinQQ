Attribute VB_Name = "TestArea"
Option Explicit

Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Ano", "Sul", "Sudeste", "Nordeste")
    ws.Range("A2:D2").Value = Array(2022, 10, 20, 5)
    ws.Range("A3:D3").Value = Array(2023, 12, 22, 6)
    ws.Range("A4:D4").Value = Array(2024, 15, 25, 8)
    ws.Range("A5:D5").Value = Array(2025, 18, 30, 9)
    Set Fixture = ChartData_FromRange(ws.Range("A1:D5"))
End Function

Private Function ComboFixture(wb As Workbook, Optional lineBig As Boolean = False) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Ano", "Vendas", "Custo", "Margem")
    ws.Range("A2:D2").Value = Array(2022, 100, 60, IIf(lineBig, 100, 0.42))
    ws.Range("A3:D3").Value = Array(2023, 120, 70, IIf(lineBig, 120, 0.45))
    ws.Range("A4:D4").Value = Array(2024, 150, 80, IIf(lineBig, 150, 0.47))
    Set ComboFixture = ChartData_FromRange(ws.Range("A1:D4"))
End Function

Public Sub Test_Area_Block_ValMidTotals()
    Dim wb As Workbook, blk As Range, body As Range
    Set wb = NewTempWorkbook()
    Set blk = Area_BuildBlock(wb, "ar0001", Fixture(wb), "area"): Set body = Block_Body(blk)
    AssertEqual 11, blk.Columns.Count: AssertEqual 3, Area_SeriesCount(blk)
    AssertEqual 35, body.Cells(2, 2).Value: AssertEqual 35, body.Cells(2, 3).Value
    AssertEqual 10, body.Cells(2, Area_ValCol(blk, 1)).Value: AssertEqual 5, body.Cells(2, Area_ValCol(blk, 3)).Value
    AssertEqual 5, body.Cells(2, Area_MidCol(blk, 1)).Value: AssertEqual 20, body.Cells(2, Area_MidCol(blk, 2)).Value
    AssertAlmost 32.5, body.Cells(2, Area_MidCol(blk, 3)).Value
    AssertEqual 3.5, body.Cells(2, Area_PosCol(blk)).Value
    AssertEqual 45, body.Cells(2, Area_VLineCol(blk)).Value, "media de 35, 40, 48, 57"
    Set blk = Area_BuildBlock(wb, "ar0002", Fixture(wb), "area100"): Set body = Block_Body(blk)
    AssertEqual 1, body.Cells(2, 3).Value
    AssertAlmost 10 / 35, body.Cells(2, Area_ValCol(blk, 1)).Value
    AssertAlmost 5 / 35, body.Cells(2, Area_MidCol(blk, 1)).Value
    AssertTrue IsError(body.Cells(2, Area_VLineCol(blk)).Value)
    CloseTemp wb
End Sub

Public Sub Test_Area_Insert_LabelsTotalsLegendBorder()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Area_Insert(d.Source.Worksheet, d, "area").Chart
    AssertEqual "area", Registry_ChartType(cht)
    AssertEqual xlAreaStacked, cht.SeriesCollection(1).ChartType
    AssertEqual RGB(255, 255, 255), cht.SeriesCollection(1).Format.Line.ForeColor.RGB, "borda branca separa as faixas"
    AssertEqual "10", Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text
    AssertEqual "5", Chart_FindSeries(cht, "tqq_lbl_3").Points(1).DataLabel.Text
    ' desvio do brief: o rotulo centrado no PONTO da primeira/ultima categoria ficaria meio fora da faixa (texto claro sobre
    ' fundo branco some): nas pontas ele vai para dentro (direita do primeiro ponto, esquerda do ultimo)
    AssertEqual xlLabelPositionCenter, Chart_FindSeries(cht, "tqq_lbl_1").Points(2).DataLabel.Position
    AssertEqual xlLabelPositionRight, Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Position
    AssertEqual xlLabelPositionLeft, Chart_FindSeries(cht, "tqq_lbl_3").Points(4).DataLabel.Position
    AssertEqual "35", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    AssertEqual xlLabelPositionAbove, Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Position
    AssertEqual 3, cht.Legend.LegendEntries.Count
    AssertEqual "right", Registry_GetOption(cht, "legend"): AssertEqual "4", Registry_GetOption(cht, "n")
    Area_SetLabels cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl_1") Is Nothing: AssertEqual "0", Registry_GetOption(cht, "labels")
    Area_SetTotals cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertEqual 3, cht.SeriesCollection.Count
    CloseTemp wb
End Sub

Public Sub Test_Area100_LabelsShowShare()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Area_Insert(d.Source.Worksheet, d, "area100").Chart
    AssertEqual xlAreaStacked100, cht.SeriesCollection(1).ChartType
    AssertEqual "29%", Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text, "10/35"
    AssertEqual "35", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "total absoluto"
    CloseTemp wb
End Sub

Public Sub Test_Area_Annotations_And_IndependentGeometry()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ln As Shape, expectStart As Double, il As Double, iw As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Area_Insert(d.Source.Worksheet, d, "area").Chart
    Set ln = Ann_AddCagr(cht, 1, 4)
    AssertEqual "CAGR +17,7%", cht.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text      ' (57/35)^(1/3)-1 = 17,65%
    Ann_AddDifference cht, 1, 4
    AssertEqual "+22 (+63%)", cht.Shapes("tqq_ann_diff_1_lbl").TextFrame2.TextRange.Text
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth
    If cht.Axes(xlCategory).AxisBetweenCategories Then expectStart = il + iw / 8 Else expectStart = il
    AssertAlmost expectStart, ln.Left, 1, "inicio da seta na categoria 1 conforme a posicao do eixo"
    CloseTemp wb
End Sub

Public Sub Test_Combo_Insert_LineSecondaryTotals()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ln As Series
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb)
    Set cht = Combo_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "combo", Registry_ChartType(cht)
    AssertEqual xlColumnStacked, cht.SeriesCollection(1).ChartType
    Set ln = Combo_LineSeries(cht)
    AssertEqual "Margem", ln.Name
    AssertEqual xlLineMarkers, ln.ChartType
    AssertEqual xlSecondary, ln.AxisGroup, "magnitudes diferem por mais de 5x: eixo secundario por padrao"
    AssertEqual "1", Registry_GetOption(cht, "secondary")
    AssertEqual xlTickLabelPositionNone, cht.Axes(xlValue, xlSecondary).TickLabelPosition
    AssertEqual "160", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "total so das colunas"
    AssertEqual 3, cht.Legend.LegendEntries.Count
    Combo_SetSecondary cht, False
    AssertEqual xlPrimary, Combo_LineSeries(cht).AxisGroup
    AssertEqual "0", Registry_GetOption(cht, "secondary")
    Ann_AddCagr cht, 1, 3
    AssertEqual "CAGR +19,9%", cht.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text      ' (230/160)^(1/2)-1
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb, True)
    Set cht = Combo_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "0", Registry_GetOption(cht, "secondary"), "mesma ordem de grandeza: eixo unico"
    AssertEqual xlPrimary, Combo_LineSeries(cht).AxisGroup
    CloseTemp wb
End Sub

Public Sub Test_Combo_RequiresTwoValueColumns()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, en As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Ano", "Valor"): ws.Range("A2:B2").Value = Array(2022, 1): ws.Range("A3:B3").Value = Array(2023, 2)
    Set d = ChartData_FromRange(ws.Range("A1:B3"))
    On Error Resume Next
    Combo_Insert ws, d
    en = Err.Number
    On Error GoTo 0
    AssertEqual vbObjectError + 6008, en
    AssertEqual 0, ws.ChartObjects.Count
    ws.Range("A1:B3").Select
    Core_ClearLastNotice
    Core_Insert "combo"
    AssertTrue InStr(Core_LastNotice(), "pelo menos 2 colunas") > 0, Core_LastNotice()
    Watchers_Stop
    CloseTemp wb
End Sub

Public Sub Test_Legend_Modes_LineEnd_Only_LineAndCombo()
    Dim wb As Workbook, d As clsChartData, cht As Chart, p As DataLabel, en As Long, ed As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "line").Chart
    Deco_SetLegend cht, "lineend"
    AssertFalse cht.HasLegend
    Set p = cht.SeriesCollection(1).Points(4).DataLabel
    AssertTrue p.ShowSeriesName, "nome da serie no ultimo ponto"
    AssertTrue InStr(p.Text, "Sul") > 0, p.Text
    AssertEqual "lineend", Registry_GetOption(cht, "legend")
    Stacked_SetLabels cht, False
    AssertTrue cht.SeriesCollection(1).Points(4).DataLabel.ShowSeriesName, "fim da linha sobrevive ao toggle de rotulos"
    Deco_SetLegend cht, "none"
    AssertFalse cht.HasLegend
    AssertFalse cht.SeriesCollection(1).Points(4).HasDataLabel, "sem rotulos e sem fim da linha"
    Deco_SetLegend cht, "right"
    AssertTrue cht.HasLegend: AssertEqual xlLegendPositionRight, cht.Legend.Position
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    On Error Resume Next
    Deco_SetLegend cht, "lineend"
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    AssertEqual "Rótulos no fim da linha só existem em gráficos de linha e combinação.", ed
    AssertEqual "right", Registry_GetOption(cht, "legend", "right")
    CloseTemp wb
End Sub

' ---------- Tarefa 6: geometria independente, carry-overs e opcao de legenda ----------

' Posicao esperada (pontos do grafico) do ponto da categoria i de uma serie de linha/area, por regra de tres sobre a area
' interna e o eixo reais (nunca Geom_*): x pelas fatias (eixo "entre categorias") ou pelas bordas, y pela escala.
Private Function ExpX(cht As Chart, i As Long, n As Long) As Double
    With cht.PlotArea
        If cht.Axes(xlCategory).AxisBetweenCategories Then ExpX = .InsideLeft + .InsideWidth * (i - 0.5) / n Else ExpX = .InsideLeft + .InsideWidth * (i - 1) / (n - 1)
    End With
End Function

Private Function ExpY(cht As Chart, v As Double, Optional grp As Long = 1) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue, grp)
    With cht.PlotArea
        ExpY = .InsideTop + .InsideHeight * (ax.MaximumScale - v) / (ax.MaximumScale - ax.MinimumScale)
    End With
End Function

' Os pontos auxiliares (rotulo no meio de cada faixa, total no topo da pilha) caem onde a pilha desenhada pede: o meio da
' faixa j da categoria i = soma das faixas anteriores + metade da faixa, lido da TABELA (nao do bloco).
Private Sub AssertAreaHelpersAligned(cht As Chart, ws As Worksheet, is100 As Boolean, tag As String)
    Dim i As Long, j As Long, cum As Double, v As Double, tot As Double, s As Series
    For i = 1 To 4
        tot = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(i + 1, 2), ws.Cells(i + 1, 4)))
        cum = 0
        For j = 1 To 3
            v = ws.Cells(i + 1, 1 + j).Value
            If is100 Then v = v / tot
            Set s = Chart_FindSeries(cht, "tqq_lbl_" & j)
            AssertAlmost ExpX(cht, i, 4), s.Points(i).Left, 0.8, tag & ": x do rotulo " & j & " na categoria " & i
            AssertAlmost ExpY(cht, cum + v / 2), s.Points(i).Top, 0.8, tag & ": y do meio da faixa " & j & " na categoria " & i
            cum = cum + v
        Next
        Set s = Chart_FindSeries(cht, "tqq_total")
        AssertAlmost ExpX(cht, i, 4), s.Points(i).Left, 0.8, tag & ": x do total " & i
        AssertAlmost ExpY(cht, IIf(is100, 1, tot)), s.Points(i).Top, 0.8, tag & ": y do total " & i & " (topo da pilha)"
    Next
End Sub

Public Sub Test_Area_Geometry_HelpersFollowBandsAndTotals_AfterDataChange()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, kind As Variant
    For Each kind In Array("area", "area100")
        Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
        Set cht = Area_Insert(ws, d, CStr(kind)).Chart
        AssertAreaHelpersAligned cht, ws, (kind = "area100"), kind & " inicial"
        ws.Range("C3").Value = 60: ws.Range("D5").Value = 40      ' muda a escala do eixo (area) e as participacoes (area100)
        cht.Refresh
        AssertAreaHelpersAligned cht, ws, (kind = "area100"), kind & " apos mudar os dados"
        cht.Axes(xlCategory).AxisBetweenCategories = False       ' areas nas bordas: as auxiliares acompanham o eixo
        cht.Refresh
        AssertAreaHelpersAligned cht, ws, (kind = "area100"), kind & " eixo sobre as marcas"
        CloseTemp wb
    Next
End Sub

' Valores nao numericos e zero: a faixa vale 0 (sem #N/D no bloco), o rotulo some pelo formato (zero vazio) e o total ignora texto e erro.
' Area 100%: total 0 nao divide por zero. (Os rotulos das auxiliares de area nao dependem de NA() para sumir: pontos NA() nao
' escondem rotulos no Excel; o zero com formato de secao vazia, sim.)
Public Sub Test_Area_NonNumericZeroAndTotalZero_NoErrorsLabelsHidden()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, blk As Range, body As Range, kind As Variant, j As Long, i As Long
    For Each kind In Array("area", "area100")
        Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
        ws.Range("A1:D1").Value = Array("Ano", "Sul", "Sudeste", "Nordeste")
        ws.Range("A2:D2").Value = Array(2022, "n/d", 20, 5)
        ws.Range("A3:D3").Value = Array(2023, 0, 0, 0)
        ws.Range("A4:D4").Value = Array(2024, 15, 25, 8)
        ws.Range("C4").Formula = "=NA()"
        Set d = ChartData_FromRange(ws.Range("A1:D4"))
        Set cht = Area_Insert(ws, d, CStr(kind)).Chart
        Set blk = Registry_Block(cht): Set body = Block_Body(blk)
        For i = 1 To 3
            For j = 1 To 3
                AssertFalse IsError(body.Cells(i + 1, Area_ValCol(blk, j)).Value), kind & ": Val sem erro (linha " & i & ", serie " & j & ")"
                AssertFalse IsError(body.Cells(i + 1, Area_MidCol(blk, j)).Value), kind & ": Mid sem erro (linha " & i & ", serie " & j & ")"
            Next
            AssertFalse IsError(body.Cells(i + 1, 2).Value), kind & ": Total sem erro (linha " & i & ")"
        Next
        AssertEqual 25, body.Cells(2, 2).Value, kind & ": o total ignora o texto"
        AssertEqual 0, body.Cells(2, Area_ValCol(blk, 1)).Value, kind & ": texto vira 0"
        AssertEqual 0, body.Cells(3, 2).Value
        AssertEqual 0, body.Cells(3, Area_ValCol(blk, 2)).Value, kind & ": total 0 sem divisao por zero"
        AssertEqual 23, body.Cells(4, 2).Value, kind & ": o total ignora o erro da origem"
        AssertEqual 0, body.Cells(4, Area_ValCol(blk, 2)).Value, kind & ": #N/D na origem vira 0"
        AssertEqual "", Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text, kind & ": rotulo do texto some"
        AssertEqual "", Chart_FindSeries(cht, "tqq_lbl_2").Points(2).DataLabel.Text, kind & ": rotulo do zero some"
        AssertEqual "", Chart_FindSeries(cht, "tqq_lbl_2").Points(3).DataLabel.Text, kind & ": rotulo do erro some"
        AssertEqual "", Chart_FindSeries(cht, "tqq_total").Points(2).DataLabel.Text, kind & ": total 0 sem texto"
        AssertFalse Core_ChartSourceMismatch(cht)
        CloseTemp wb
    Next
End Sub

' Carry-over (a): a area nao tem barras, entao a seta de diferenca fica a 10 pt do ponto da ultima categoria (sem meia largura
' de barra), com o eixo "entre categorias" e com o eixo sobre as marcas; nas colunas e na combinacao a meia barra continua.
Public Sub Test_Area_DiffArrow_HasNoBarOffset_ColumnsAndComboKeepIt()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ln As Shape, between As Variant, xTo As Double, slot As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Area_Insert(d.Source.Worksheet, d, "area").Chart
    Set ln = Ann_AddDifference(cht, 1, 4)
    AssertEqual 0, Geom_ChartBarHalfWidth(cht), "area: sem barras"
    For Each between In Array(False, True)
        cht.Axes(xlCategory).AxisBetweenCategories = CBool(between)
        cht.Refresh
        Ann_Refresh cht
        With cht.PlotArea
            If between Then xTo = .InsideLeft + .InsideWidth * 3.5 / 4 Else xTo = .InsideLeft + .InsideWidth
        End With
        AssertAlmost xTo + 10, ln.Left, 0.8, "area (entre categorias = " & between & "): seta a 10 pt do ponto da ultima categoria"
    Next
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Area_Insert(d.Source.Worksheet, d, "area100").Chart
    AssertEqual 0, Geom_ChartBarHalfWidth(cht), "area100: sem barras"
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    Set ln = Ann_AddDifference(cht, 1, 4)
    slot = cht.PlotArea.InsideWidth / 4
    AssertAlmost cht.PlotArea.InsideLeft + slot * 3.5 + slot / 1.5 / 2 + 10, ln.Left, 0.8, "colunas: 10 pt depois da borda da ultima coluna (gap 50%)"
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb)
    Set cht = Combo_Insert(d.Source.Worksheet, d).Chart
    Set ln = Ann_AddDifference(cht, 1, 3)
    slot = cht.PlotArea.InsideWidth / 3
    AssertAlmost cht.PlotArea.InsideLeft + slot * 2.5 + slot / 1.5 / 2 + 10, ln.Left, 0.8, "combinacao: como as colunas"
    CloseTemp wb
End Sub

' Carry-over (b): HideHelperLegendEntries nas pilhas de AREA com a auxiliar de totais e as de rotulo (no mesmo grupo primario
' que as areas, em grupo de linha): com a legenda a direita/esquerda a lista vem invertida (ultima serie primeiro), em cima/embaixo
' na ordem normal; o mapeamento tira so as 4 auxiliares (3 de rotulo e a de total) e deixa as 3 areas, identificadas pela cor da chave.
Public Sub Test_Chart_HideHelperLegendEntries_AreaStacks_WithHelperTotals()
    Dim wb As Workbook, d As clsChartData, cht As Chart, kind As Variant, lp As Variant, pal As Variant, k As Long, keys As String, want As String
    pal = Style_Palette("padrao")
    For Each kind In Array("area", "area100")
        Set wb = NewTempWorkbook(): Set d = Fixture(wb)
        Set cht = Area_Insert(d.Source.Worksheet, d, CStr(kind)).Chart
        AssertEqual 3, cht.Legend.LegendEntries.Count, kind & ": legenda ja nasce sem as auxiliares"
        For Each lp In Array(xlLegendPositionRight, xlLegendPositionBottom)       ' esquerda e em cima: o teste do mesmo grupo (abaixo) cobre
            cht.Legend.Position = lp
            cht.HasLegend = False: cht.HasLegend = True: cht.Legend.Position = lp      ' religa: todas as 7 entradas voltam
            AssertEqual 7, cht.Legend.LegendEntries.Count, kind & ": premissa (" & lp & ")"
            Chart_HideHelperLegendEntries cht
            AssertEqual 3, cht.Legend.LegendEntries.Count, kind & ": so as 3 areas (" & lp & ")"
            keys = ""
            For k = 1 To 3: keys = keys & cht.Legend.LegendEntries(k).LegendKey.Format.Fill.ForeColor.RGB & ";": Next
            If lp = xlLegendPositionRight Then
                want = CLng(pal(2)) & ";" & CLng(pal(1)) & ";" & CLng(pal(0)) & ";"
            Else
                want = CLng(pal(0)) & ";" & CLng(pal(1)) & ";" & CLng(pal(2)) & ";"
            End If
            AssertEqual want, keys, kind & ": entradas e ordem da legenda (" & lp & ")"
            Chart_HideHelperLegendEntries cht
            AssertEqual 3, cht.Legend.LegendEntries.Count, kind & ": idempotente (" & lp & ")"
        Next
        CloseTemp wb
    Next
End Sub

' Combinacao: as colunas empilhadas (invertidas na legenda a direita), a auxiliar de total (linha, grupo primario) e a linha
' no grupo secundario: sobram as 2 colunas e a linha, nao a auxiliar.
Public Sub Test_Chart_HideHelperLegendEntries_Combo_ColumnsAndLine()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lp As Variant, pal As Variant, k As Long, keys As String
    pal = Style_Palette("padrao")
    For Each lp In Array(xlLegendPositionRight, xlLegendPositionBottom)
        Set wb = NewTempWorkbook(): Set d = ComboFixture(wb)
        Set cht = Combo_Insert(d.Source.Worksheet, d).Chart
        cht.Legend.Position = lp
        cht.HasLegend = False: cht.HasLegend = True: cht.Legend.Position = lp
        AssertEqual 4, cht.Legend.LegendEntries.Count, "premissa (" & lp & ")"
        Chart_HideHelperLegendEntries cht
        AssertEqual 3, cht.Legend.LegendEntries.Count, "combinacao: 2 colunas e a linha (" & lp & ")"
        keys = ""
        For k = 1 To 2: keys = keys & cht.Legend.LegendEntries(k).LegendKey.Format.Fill.ForeColor.RGB & ";": Next
        If lp = xlLegendPositionRight Then
            AssertEqual CLng(pal(1)) & ";" & CLng(pal(0)) & ";", keys, "colunas invertidas na legenda a direita"
        Else
            AssertEqual CLng(pal(0)) & ";" & CLng(pal(1)) & ";", keys, "colunas na ordem normal embaixo"
        End If
        AssertEqual CLng(pal(2)), cht.Legend.LegendEntries(3).LegendKey.Format.Line.ForeColor.RGB, "a terceira entrada e a linha (" & lp & ")"
        CloseTemp wb
    Next
End Sub

' Carry-over (f): a borda das areas e uma opcao do grafico ("border", em pt; padrao 1) lida em Style_ApplySeriesColors:
' trocar a paleta mantem a espessura, e o parametro explicito (sem persistir) a sobrepoe. A cor do rotulo de cada faixa segue a
' faixa nova (Area_ApplyColors).
Public Sub Test_Area_PaletteChange_KeepsBorderWeightAndRelabelsContrast()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, pal As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Area_Insert(d.Source.Worksheet, d, "area").Chart
    AssertEqual 1, cht.SeriesCollection(1).Format.Line.Weight, "borda padrao: 1 pt"
    AssertEqual "", Registry_GetOption(cht, "border"), "o padrao nao grava a opcao"
    ApplyPaletteViaCore cht, "cinza"
    pal = Style_Palette("cinza")
    AssertEqual CLng(pal(1)), cht.SeriesCollection(2).Format.Fill.ForeColor.RGB, "paleta cinza"
    AssertEqual RGB(255, 255, 255), cht.SeriesCollection(2).Format.Line.ForeColor.RGB, "a borda continua branca"
    AssertEqual 1, cht.SeriesCollection(2).Format.Line.Weight
    AssertEqual RGB(255, 255, 255), Chart_FindSeries(cht, "tqq_lbl_1").Points(2).DataLabel.Font.Color, "faixa escura (64,64,64): rotulo claro"
    Registry_SetOption cht, "border", "1.5"
    ApplyPaletteViaCore cht, "padrao"
    AssertEqual 1.5, cht.SeriesCollection(1).Format.Line.Weight, "a espessura da opcao sobrevive a troca de paleta"
    AssertEqual 1.5, cht.SeriesCollection(3).Format.Line.Weight
    Style_ApplySeriesColors cht, "padrao", 2.25
    AssertEqual 2.25, cht.SeriesCollection(1).Format.Line.Weight, "o parametro explicito vale"
    AssertEqual "1.5", Registry_GetOption(cht, "border"), "o parametro nao persiste"
    AssertEqual RGB(255, 255, 255), cht.SeriesCollection(1).Format.Line.ForeColor.RGB
    CloseTemp wb
End Sub

Private Sub ApplyPaletteViaCore(cht As Chart, name As String)
    cht.Parent.Select
    Core_SetOption "palette", name
End Sub

' Combinacao: reversao do que Combo_Insert faz em falha depois de alocar o bloco (grafico e bloco apagados, sem nome orfao).
Public Sub Test_Combo_Insert_RollsBack_OnFailure()
    Dim wb As Workbook, d As clsChartData, en As Long, nm As Name, orphan As Boolean
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb)
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Combo_Insert d.Source.Worksheet, d
    en = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, en
    AssertEqual 0, d.Source.Worksheet.ChartObjects.Count, "grafico desfeito"
    For Each nm In wb.Names
        If Left$(nm.Name, Len(ID_PREFIX)) = ID_PREFIX Then orphan = True
    Next
    AssertFalse orphan, "nome do bloco desfeito"
    CloseTemp wb
End Sub

' O eixo secundario e a linha por regra de tres: os pontos da linha seguem a escala SECUNDARIA (0,39..0,48), as colunas e os
' totais a primaria; religar no eixo primario mexe so na linha.
Public Sub Test_Combo_Geometry_LineOnSecondaryScale_ColumnsOnPrimary()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ln As Series, tot As Series, i As Long, v As Variant, sum As Variant
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb)
    Set cht = Combo_Insert(d.Source.Worksheet, d).Chart
    Set ln = Combo_LineSeries(cht): Set tot = Chart_FindSeries(cht, "tqq_total")
    v = Array(0.42, 0.45, 0.47): sum = Array(160, 190, 230)
    AssertTrue cht.Axes(xlValue, xlSecondary).MinimumScale = 0 And cht.Axes(xlValue, xlSecondary).MaximumScale >= 0.47 And cht.Axes(xlValue, xlSecondary).MaximumScale <= 0.6, "escala secundaria propria, de 0 ao maximo da linha"
    AssertTrue cht.Axes(xlValue, xlPrimary).MaximumScale >= 230, "escala primaria dos totais"
    For i = 1 To 3
        AssertAlmost ExpX(cht, i, 3), ln.Points(i).Left, 0.8, "x da linha " & i
        AssertAlmost ExpY(cht, CDbl(v(i - 1)), xlSecondary), ln.Points(i).Top, 0.8, "y da linha na escala secundaria (" & i & ")"
        AssertAlmost ExpY(cht, CDbl(sum(i - 1)), xlPrimary), tot.Points(i).Top, 0.8, "y do total na escala primaria (" & i & ")"
    Next
    Combo_SetSecondary cht, False
    cht.Refresh
    For i = 1 To 3
        AssertAlmost ExpY(cht, CDbl(v(i - 1)), xlPrimary), ln.Points(i).Top, 0.8, "linha no eixo primario (" & i & ")"
        AssertAlmost ExpY(cht, CDbl(sum(i - 1)), xlPrimary), tot.Points(i).Top, 0.8, "o total segue na primaria"
    Next
    AssertFalse cht.HasAxis(xlValue, xlSecondary), "sem a linha no secundario o eixo some"
    Combo_SetSecondary cht, True
    AssertTrue cht.HasAxis(xlValue, xlSecondary)
    AssertEqual xlTickLabelPositionNone, cht.Axes(xlValue, xlSecondary).TickLabelPosition
    AssertEqual 3, cht.Legend.LegendEntries.Count, "legenda continua sem a auxiliar depois de trocar de eixo"
    CloseTemp wb
End Sub

' O padrao do eixo secundario e uma regra sobre os dados (5x), nos dois sentidos e com texto/vazio/zero.
Public Sub Test_Combo_DefaultSecondary_Rule()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Col", "Linha")
    ws.Range("A2:C2").Value = Array(2022, 100, 600)
    ws.Range("A3:C3").Value = Array(2023, 80, 450)
    AssertTrue Combo_DefaultSecondary(ChartData_FromRange(ws.Range("A1:C3"))), "linha 6x maior"
    ws.Range("C2").Value = 500: ws.Range("C3").Value = 400
    AssertFalse Combo_DefaultSecondary(ChartData_FromRange(ws.Range("A1:C3"))), "exatamente 5x nao conta"
    ws.Range("C2").Value = 10: ws.Range("C3").Value = 5
    AssertTrue Combo_DefaultSecondary(ChartData_FromRange(ws.Range("A1:C3"))), "linha 10x menor"
    ws.Range("C2").Value = -10: ws.Range("C3").Value = -5
    AssertTrue Combo_DefaultSecondary(ChartData_FromRange(ws.Range("A1:C3"))), "valores negativos contam em modulo"
    ws.Range("C2").Value = "n/d": ws.Range("C3").Value = Empty
    AssertFalse Combo_DefaultSecondary(ChartData_FromRange(ws.Range("A1:C3"))), "linha sem numeros: magnitude 0"
    ws.Range("A1:D1").Value = Array("Ano", "A", "B", "Linha")
    ws.Range("A2:D2").Value = Array(2022, 100, -40, 500)
    ws.Range("A3:D3").Value = Array(2023, 100, -50, 350)
    AssertTrue Combo_DefaultSecondary(ChartData_FromRange(ws.Range("A1:D3"))), "a soma das colunas e liquida: 60 e 50 contra 500 (8x)"
    ws.Range("C2:C3").ClearContents: ws.Range("B2:B3").ClearContents
    AssertFalse Combo_DefaultSecondary(ChartData_FromRange(ws.Range("A1:D3"))), "colunas sem numeros: magnitude 0"
    CloseTemp wb
End Sub

' A linha da combinacao mostra o formato da CELULA (o numfmt e um so para o grafico e valeria para as colunas e os totais): com o
' numfmt int o 0,42 apareceria como "0". Mudar o numfmt muda as colunas e os totais, nao a linha; mudar a celula muda a linha.
' (Celula em Geral: formato de reserva, ver Test_Combo_LineLabelFormat_GeneralCellFallbackRule_AndFollowsCellAfterwards.)
Public Sub Test_Combo_LineLabels_FollowSourceCellFormat_NumfmtKeepsColumnsAndTotals()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ln As Series, dec As String
    dec = Application.International(xlDecimalSeparator)
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb)
    Set cht = Combo_Insert(d.Source.Worksheet, d).Chart
    Set ln = Combo_LineSeries(cht)
    AssertEqual "42" & dec & "0%", ln.Points(1).DataLabel.Text, "celula geral com |v| <= 1: formato de reserva pct1 (antes 0,42)"
    d.Source.Worksheet.Range("D2:D4").NumberFormat = "0%"
    Core_SafeRefresh cht
    AssertEqual "42%", ln.Points(1).DataLabel.Text, "celula em percentual: 42%"
    cht.Parent.Select
    Core_SetOption "numfmt", "dec1"
    AssertEqual "100" & dec & "0", cht.SeriesCollection(1).Points(1).DataLabel.Text, "coluna: numfmt dec1"
    AssertEqual "160" & dec & "0", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "total: numfmt dec1"
    AssertEqual "42%", ln.Points(1).DataLabel.Text, "linha: segue a celula"
    Watchers_Stop
    CloseTemp wb
End Sub

' A API de opcoes numa combinacao: toggles de totais, rotulos, conectores (so nas colunas) e eixo secundario; gap das colunas;
' paleta e formato; e o "fim da linha" da combinacao (a linha ganha o nome no ultimo ponto, sem legenda).
Public Sub Test_Core_Combo_TogglesOptionsAndLineEnd()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, ln As Series, p As DataLabel
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:D4").Select
    Core_Insert "combo"
    Set cht = Core_ActiveChart()
    Set ln = Combo_LineSeries(cht)
    AssertEqual "right", Core_GetOption("legend")
    AssertTrue Core_GetToggle("secondary"): AssertTrue Core_GetToggle("totals"): AssertTrue Core_GetToggle("labels")
    Core_SetToggle "secondary", False
    AssertEqual xlPrimary, ln.AxisGroup: AssertFalse Core_GetToggle("secondary")
    Core_SetToggle "secondary", True
    AssertEqual xlSecondary, ln.AxisGroup: AssertTrue Core_GetToggle("secondary")
    Core_SetToggle "totals", False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    Core_SetToggle "totals", True
    AssertEqual "160", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "totais religados: so as colunas"
    Core_SetToggle "connectors", True
    AssertTrue cht.ChartGroups(1).HasSeriesLines, "conectores nas colunas"
    AssertEqual "1", Core_GetOption("connectors")
    Core_SetOption "gap", "120"
    AssertEqual 120, cht.ChartGroups(1).GapWidth
    Core_SetToggle "labels", False
    AssertFalse cht.SeriesCollection(1).HasDataLabels: AssertFalse ln.HasDataLabels
    Core_SetToggle "labels", True
    AssertTrue cht.SeriesCollection(1).HasDataLabels: AssertTrue ln.HasDataLabels
    Core_SetOption "legend", "lineend"
    AssertFalse cht.HasLegend: AssertEqual "lineend", Core_GetOption("legend")
    Set p = ln.Points(3).DataLabel
    AssertTrue p.ShowSeriesName: AssertTrue InStr(p.Text, "Margem") > 0, p.Text
    Core_SetToggle "labels", False
    AssertTrue ln.Points(3).DataLabel.ShowSeriesName, "o nome no fim da linha sobrevive a desligar os rotulos"
    AssertFalse ln.Points(3).DataLabel.ShowValue
    AssertFalse ln.Points(2).HasDataLabel, "so o ultimo ponto"
    Core_SetOption "legend", "right"
    AssertTrue cht.HasLegend
    AssertFalse ln.Points(3).HasDataLabel, "sem rotulos e sem fim da linha"
    AssertEqual 3, cht.Legend.LegendEntries.Count, "a legenda volta sem a auxiliar"
    Watchers_Stop
    CloseTemp wb
End Sub

' Linhas de area: "gap" e "secondary" nao se aplicam e avisam sem falha; os toggles de totais/rotulos da area passam pelo Core.
Public Sub Test_Core_Area_TogglesOptionsAndRefusals()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:D5").Select
    Core_Insert "area"
    Set cht = Core_ActiveChart()
    AssertEqual "area", Registry_ChartType(cht)
    Core_ClearLastNotice
    Core_SetOption "gap", "100"
    AssertTrue InStr(Core_LastNotice(), "não se aplica") > 0, Core_LastNotice()
    AssertEqual "", Core_GetOption("gap")
    Core_SetToggle "secondary", True
    AssertFalse Core_GetToggle("secondary")
    Core_SetToggle "totals", False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    Core_SetToggle "totals", True
    AssertEqual "35", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    Core_SetToggle "labels", False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl_1") Is Nothing
    Core_SetToggle "labels", True
    AssertEqual "10", Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text
    Core_ClearLastNotice
    Core_SetToggle "connectors", True
    AssertEqual "'Conectores' não se aplica a este tipo de gráfico.", Core_LastNotice(), "area nao tem conectores: aviso (Tarefa 11)"
    AssertEqual "", Core_GetOption("connectors"), "area nao tem conectores: a opcao nem e gravada"
    Core_SetOption "numfmt", "dec1"
    AssertEqual "10" & Application.International(xlDecimalSeparator) & "0", Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text, "numfmt chega aos rotulos de segmento"
    AssertEqual "35" & Application.International(xlDecimalSeparator) & "0", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "e aos totais"
    Core_SetOption "legend", "lineend"
    AssertTrue InStr(Core_LastNotice(), "Rótulos no fim da linha") > 0
    AssertEqual "right", Core_GetOption("legend")
    Core_SetOption "legend", "none"
    AssertFalse cht.HasLegend
    Core_SetOption "legend", "right"
    AssertEqual 3, cht.Legend.LegendEntries.Count, "legenda ligada de novo, sem as auxiliares"
    Watchers_Stop
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:D5").Select
    Core_Insert "area100"
    Set cht = Core_ActiveChart()
    Core_SetOption "numfmt", "pct1"
    AssertEqual "28" & Application.International(xlDecimalSeparator) & "6%", Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text, "area100: participacao com 1 casa"
    AssertEqual "35", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "o total absoluto nao vira percentual"
    Watchers_Stop
    CloseTemp wb
End Sub

' A legenda ligada depois que a area de plotagem foi fixada pela folga das setas/totais nao avanca sobre ela.
Public Sub Test_Legend_Right_AfterManualPlotLayout_PlotStaysLeftOfLegend()
    Dim wb As Workbook, d As clsChartData, cht As Chart, kind As Variant
    For Each kind In Array("area", "stacked")
        Set wb = NewTempWorkbook(): Set d = Fixture(wb)
        If kind = "area" Then Set cht = Area_Insert(d.Source.Worksheet, d, "area").Chart Else Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
        Deco_SetLegend cht, "none"
        cht.Refresh
        AssertTrue cht.PlotArea.Width > cht.ChartArea.Width - 30, kind & ": sem legenda a area de plotagem ocupa o grafico"
        Ann_AddCagr cht, 1, 4                                  ' fixa o tamanho da area de plotagem (layout manual)
        Deco_SetLegend cht, "right"
        cht.Refresh
        AssertTrue cht.HasLegend: AssertEqual xlLegendPositionRight, cht.Legend.Position
        AssertTrue cht.PlotArea.Left + cht.PlotArea.Width <= cht.Legend.Left - 9, kind & ": area de plotagem antes da legenda (direita da area=" & (cht.PlotArea.Left + cht.PlotArea.Width) & ", legenda=" & cht.Legend.Left & ")"
        CloseTemp wb
    Next
End Sub

' Barras com seta de CAGR: ligar a legenda depois refaz a folga (a ponta da seta fica a 10 pt do fim da barra e o rotulo cabe entre
' ela e a legenda, sem o limite que empurra o rotulo para cima da ponta) em vez de a folga so valer na criacao. Sem totais, para a
' ponta ser so "fim da barra + 10 pt" (regra de tres sobre a area interna e o eixo reais).
Public Sub Test_Legend_Bars_Cagr_RightLegendAfterNone_ArrowAndLabelLeftOfLegend()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ln As Shape, lbl As Shape, ax As Axis, xEnd As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Stacked_SetTotals cht, False
    Deco_SetLegend cht, "none"
    Set ln = Ann_AddCagr(cht, 1, 4)
    Set lbl = cht.Shapes("tqq_ann_cagr_1_lbl")
    Deco_SetLegend cht, "right"
    Core_SafeRefresh cht
    cht.Refresh
    AssertTrue cht.HasLegend
    Set ax = cht.Axes(xlValue)
    xEnd = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (57 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)      ' 2025: 18 + 30 + 9
    AssertAlmost xEnd + 10, ln.Left + ln.Width, 0.8, "ponta da seta a 10 pt do fim da barra (sem ser empurrada pelo limite)"
    AssertTrue lbl.Left >= ln.Left + ln.Width + 2, "rotulo a direita da ponta (rotulo L=" & lbl.Left & ", ponta=" & (ln.Left + ln.Width) & ")"
    AssertTrue lbl.Left + lbl.Width <= cht.Legend.Left + 0.5, "rotulo do CAGR antes da legenda (direita do rotulo=" & (lbl.Left + lbl.Width) & ", legenda=" & cht.Legend.Left & ")"
    CloseTemp wb
End Sub

' Largura MEDIDA (pt) de um rotulo de 9 pt em negrito com so digitos e separadores (Calibri: margem 6,0, digito 4,56, separador
' 2,4; ver o comentario de Labels_EstimateWidth), independente das constantes do estimador de producao (que superestimam).
Private Function MeasuredNumberWidth(txt As String) As Double
    Dim i As Long, w As Double, ch As String
    w = 6
    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch >= "0" And ch <= "9" Then w = w + 4.56 Else w = w + 2.4
    Next
    MeasuredNumberWidth = w
End Function

' Mudar o formato numerico alarga os rotulos de total das barras 100% (que ficam no fim do eixo, junto da legenda): a folga dos
' totais e refeita (carry-over da Tarefa 3). "35" vira "35,00": mais ~11,5 pt que a folga de criacao (estimada para "35") nao cobre.
' O ponto XY do total fica no fim do eixo (Point.Left, confiavel) e o rotulo comeca ~5,5 pt depois dele; a geometria do proprio
' DataLabel nao e lida (no Excel invisivel dos testes ela fica na posicao da primeira leitura).
Public Sub Test_Core_Numfmt_Bars100_TotalLabelStillLeftOfLegendAfterWiderFormat()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, tot As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:D5").Select
    Core_Insert "barstacked100"
    Set cht = Core_ActiveChart()
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertTrue tot.Points(1).Left + 5.5 + MeasuredNumberWidth(tot.Points(1).DataLabel.Text) <= cht.Legend.Left, "premissa: antes de mudar o formato o rotulo cabe"
    Core_SetOption "numfmt", "dec2"
    cht.Refresh
    AssertEqual "35" & Application.International(xlDecimalSeparator) & "00", tot.Points(1).DataLabel.Text, "premissa: o rotulo ficou mais largo"
    AssertTrue tot.Points(1).Left + 5.5 + MeasuredNumberWidth(tot.Points(1).DataLabel.Text) <= cht.Legend.Left, "rotulo do total mais largo ainda antes da legenda (ponto=" & tot.Points(1).Left & ", legenda=" & cht.Legend.Left & ")"
    Watchers_Stop
    CloseTemp wb
End Sub

' "Fim da linha": o rotulo "nome valor" do ultimo ponto cabe antes da borda do grafico mesmo com nome comprido (a area de plotagem
' sem legenda se expande e o ultimo ponto fica a meia fatia da borda).
Public Sub Test_Legend_LineEnd_LongSeriesName_LabelStaysInsideChart()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, lbl As DataLabel
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Regiao Sudeste e Centro-Oeste", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 100, 50): ws.Range("A3:C3").Value = Array(2023, 120, 60): ws.Range("A4:C4").Value = Array(2024, 140, 80)
    Set d = ChartData_FromRange(ws.Range("A1:C4"))
    Set cht = Stacked_Insert(ws, d, "line").Chart
    Deco_SetLegend cht, "lineend"
    cht.Refresh
    AssertLineEndLabelFits cht, "nome longo"
    CloseTemp wb
End Sub


' A opcao de legenda e uma ESCOLHA: "Alterar dados" a reaplica (nao a sobrescreve pelo estado real do grafico novo).
Public Sub Test_Core_RebuildFromRange_ReappliesSavedLegendAndSecondaryChoices()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, p As DataLabel
    ' coluna empilhada: sem legenda
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Stacked_Insert(ws, d, "stacked").Chart
    Deco_SetLegend cht, "none"
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:D5"))
    Set cht = Core_ActiveChart()
    AssertFalse cht.HasLegend, "a escolha 'nenhuma' volta a valer"
    AssertEqual "none", Registry_GetOption(cht, "legend")
    Deco_SetLegend cht, "right"
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:D5"))
    Set cht = Core_ActiveChart()
    AssertTrue cht.HasLegend: AssertEqual "right", Registry_GetOption(cht, "legend")
    AssertEqual 3, cht.Legend.LegendEntries.Count
    Watchers_Stop
    CloseTemp wb
    ' linha: nome no fim, com e sem rotulos
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Stacked_Insert(ws, d, "line").Chart
    Deco_SetLegend cht, "lineend"
    Stacked_SetLabels cht, False
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:D5"))
    Set cht = Core_ActiveChart()
    AssertFalse cht.HasLegend
    AssertEqual "lineend", Registry_GetOption(cht, "legend")
    Set p = cht.SeriesCollection(2).Points(4).DataLabel
    AssertTrue p.ShowSeriesName, "nome no fim da linha reaplicado": AssertFalse p.ShowValue, "rotulos desligados continuam desligados"
    AssertFalse cht.SeriesCollection(2).Points(3).HasDataLabel
    Stacked_SetLabels cht, True
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:D5"))
    Set cht = Core_ActiveChart()
    Set p = cht.SeriesCollection(2).Points(4).DataLabel
    AssertTrue p.ShowSeriesName: AssertTrue p.ShowValue: AssertTrue InStr(p.Text, "Sudeste") > 0, p.Text
    Watchers_Stop
    CloseTemp wb
    ' combinacao: eixo secundario escolhido (nao o padrao pelos dados) e legenda
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb): Set ws = d.Source.Worksheet
    Set cht = Combo_Insert(ws, d).Chart
    AssertEqual "1", Registry_GetOption(cht, "secondary"): AssertEqual xlSecondary, Combo_LineSeries(cht).AxisGroup
    Combo_SetSecondary cht, False
    Deco_SetLegend cht, "none"
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:D4"))
    Set cht = Core_ActiveChart()
    AssertEqual "combo", Registry_ChartType(cht)
    AssertEqual "0", Registry_GetOption(cht, "secondary"): AssertEqual xlPrimary, Combo_LineSeries(cht).AxisGroup, "a escolha do usuario ganha do padrao"
    AssertFalse cht.HasLegend
    Combo_SetSecondary cht, True
    Deco_SetLegend cht, "lineend"
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:D4"))
    Set cht = Core_ActiveChart()
    AssertEqual xlSecondary, Combo_LineSeries(cht).AxisGroup
    AssertTrue Combo_LineSeries(cht).Points(3).DataLabel.ShowSeriesName
    Watchers_Stop
    CloseTemp wb
End Sub

' "Alterar dados" de uma combinacao para uma tabela de 1 coluna de valores e uma recusa de validacao (6008): aviso, o grafico antigo fica.
Public Sub Test_Core_RebuildFromRange_ComboNeedsTwoColumns_NoticeAndKeepsOld()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb): Set ws = d.Source.Worksheet
    Set cht = Combo_Insert(ws, d).Chart
    Core_ClearLastNotice
    AssertFalse Core_RebuildFromRange(cht, ws.Range("A1:B4"))
    AssertTrue InStr(Core_LastNotice(), "pelo menos 2 colunas") > 0, Core_LastNotice()
    AssertEqual 1, ws.ChartObjects.Count, "o grafico antigo fica"
    AssertEqual "combo", Registry_ChartType(cht)
    CloseTemp wb
End Sub

' "Alterar dados" de uma area: mantem o tipo, rotulos, totais, formato, paleta, borda e a legenda escolhida; troca o numero de series.
Public Sub Test_Core_RebuildFromRange_Area_KeepsKindOptionsAndLegend()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, kind As Variant
    For Each kind In Array("area", "area100")
        Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
        Set cht = Area_Insert(ws, d, CStr(kind)).Chart
        Registry_SetOption cht, "border", "1.5"
        cht.Parent.Select
        Core_SetOption "numfmt", IIf(kind = "area", "dec1", "pct1")
        Core_SetOption "palette", "cinza"
        Core_SetToggle "totals", False
        Core_SetOption "legend", "none"
        AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:C5"))                 ' 2 series em vez de 3
        Set cht = Core_ActiveChart()
        AssertEqual CStr(kind), Registry_ChartType(cht)
        AssertEqual 2, Area_SeriesCount(Registry_Block(cht))
        AssertEqual 4, cht.SeriesCollection.Count, kind & ": 2 areas + 2 rotulos (sem o total)"
        AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing, kind & ": totais continuam desligados"
        AssertFalse cht.HasLegend: AssertEqual "none", Registry_GetOption(cht, "legend")
        AssertEqual "cinza", Registry_GetOption(cht, "palette")
        AssertEqual 1.5, cht.SeriesCollection(1).Format.Line.Weight, kind & ": a espessura da borda copiada"
        AssertEqual CLng(Style_Palette("cinza")(1)), cht.SeriesCollection(2).Format.Fill.ForeColor.RGB
        AssertEqual IIf(kind = "area", "10" & Application.International(xlDecimalSeparator) & "0", "33" & Application.International(xlDecimalSeparator) & "3%"), Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text
        AssertFalse Core_ChartSourceMismatch(cht)
        Watchers_Stop
        CloseTemp wb
    Next
End Sub

' "Fim da linha": mudar o formato (rotulo "nome valor" mais largo) ou o tamanho do grafico refaz a folga do rotulo do ultimo ponto.
Public Sub Test_Legend_LineEnd_NumfmtAndResize_KeepLabelInsideChart()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, co As ChartObject
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Regiao Sudeste", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 1000000, 50): ws.Range("A3:C3").Value = Array(2023, 1200000, 60): ws.Range("A4:C4").Value = Array(2024, 1400000, 80)
    Set d = ChartData_FromRange(ws.Range("A1:C4"))
    Set co = Stacked_Insert(ws, d, "line"): Set cht = co.Chart
    cht.Parent.Select
    Core_SetOption "legend", "lineend"
    Core_SetOption "numfmt", "dec2"
    cht.Refresh
    AssertLineEndLabelFits cht, "apos o numfmt dec2 (rotulo mais largo)"
    co.Width = 300
    Chart_EnsureRoomAfterChange cht
    cht.Refresh
    AssertLineEndLabelFits cht, "apos reduzir o grafico para 300 pt"
    Watchers_Stop
    CloseTemp wb
End Sub

' Carry-over (b), o ramo que importa: auxiliares tqq_ no MESMO grupo empilhado das areas (como na cascata/Mekko, intercaladas:
' S1, tqq_a, S2, tqq_b). Com a legenda a direita/esquerda a lista vem invertida (ultima serie primeiro); em cima/embaixo na ordem
' normal. O mapeamento tem de apagar so as entradas tqq_ nos dois casos, em area empilhada e 100%. (Nas areas do ThinQQ as
' auxiliares sao linhas, em outro grupo; este e o teste do ramo xlAreaStacked* de LegendGroupReversed.)
Public Sub Test_Chart_HideHelperLegendEntries_AreaStacks_SameGroupAuxiliaries_ReversedAndNot()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, cht As Chart, s As Series, j As Long, lp As Variant, ct As Variant, k As Long, keys As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:E1").Value = Array("Cat", "S1", "tqq_a", "S2", "tqq_b")
    ws.Range("A2:E2").Value = Array("A", 10, 1, 5, 2)
    ws.Range("A3:E3").Value = Array("B", 20, 1, 6, 2)
    For Each ct In Array(xlAreaStacked, xlAreaStacked100)
        For Each lp In Array(xlLegendPositionRight, xlLegendPositionLeft, xlLegendPositionBottom, xlLegendPositionTop)
            Set co = ws.ChartObjects.Add(10, 10, 460, 280): Set cht = co.Chart
            Do While cht.SeriesCollection.Count > 0: cht.SeriesCollection(1).Delete: Loop
            For j = 1 To 4
                Set s = cht.SeriesCollection.NewSeries
                s.Name = "=Dados!" & ws.Cells(1, 1 + j).Address
                s.Values = "=Dados!" & ws.Range(ws.Cells(2, 1 + j), ws.Cells(3, 1 + j)).Address
                s.XValues = "=Dados!$A$2:$A$3"
            Next
            cht.ChartType = CLng(ct)
            For j = 1 To 4: cht.SeriesCollection(j).Format.Fill.ForeColor.RGB = RGB(10 * j, 0, 0): Next
            cht.HasLegend = True: cht.Legend.Position = lp
            AssertEqual 4, cht.Legend.LegendEntries.Count
            Chart_HideHelperLegendEntries cht
            AssertEqual 2, cht.Legend.LegendEntries.Count, "so S1 e S2 (area " & ct & ", legenda " & lp & ")"
            keys = ";"
            For k = 1 To 2: keys = keys & (cht.Legend.LegendEntries(k).LegendKey.Format.Fill.ForeColor.RGB Mod 256) & ";": Next
            AssertTrue InStr(keys, ";10;") > 0 And InStr(keys, ";30;") > 0, "entradas restantes = S1 (10) e S2 (30); chaves " & keys & " (area " & ct & ", legenda " & lp & ")"
            Chart_HideHelperLegendEntries cht
            AssertEqual 2, cht.Legend.LegendEntries.Count, "idempotente"
            co.Delete
        Next
    Next
    CloseTemp wb
End Sub

' "Fim da linha" com os rotulos desligados reserva folga so para o nome; ligar os rotulos depois acrescenta o valor ao texto do
' ultimo ponto e a folga cresce junto (Deco_ApplyLineEnd, chamada por Stacked_SetLabels).
Public Sub Test_Legend_LineEnd_LabelsToggledOn_MakesRoomForValue()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Regiao Sudeste", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 1000000, 50): ws.Range("A3:C3").Value = Array(2023, 1200000, 60): ws.Range("A4:C4").Value = Array(2024, 1400000, 80)
    Set d = ChartData_FromRange(ws.Range("A1:C4"))
    Set cht = Stacked_Insert(ws, d, "line").Chart
    Stacked_SetLabels cht, False
    Deco_SetLegend cht, "lineend"
    AssertEqual "Regiao Sudeste", cht.SeriesCollection(1).Points(3).DataLabel.Text, "sem rotulos: so o nome"
    AssertLineEndLabelFits cht, "so o nome"
    Stacked_SetLabels cht, True
    AssertEqual "Regiao Sudeste 1" & Application.International(xlThousandsSeparator) & "400" & Application.International(xlThousandsSeparator) & "000", cht.SeriesCollection(1).Points(3).DataLabel.Text, "com rotulos: nome e valor"
    AssertLineEndLabelFits cht, "nome e valor"
    CloseTemp wb
End Sub

' ---------- Rodada de correcao 1: legenda devolve o espaco, categorias de data, rollback, formato da linha, areas negativas ----------

' Largura (pt) de um rotulo de dados de 9 pt MEDIDA numa caixa de texto de ajuste automatico (margens de 3 pt, a margem interna do
' rotulo) no proprio grafico: oraculo independente do estimador de producao (Labels_EstimateWidth, que superestima ~30%). Conferido
' contra DataLabel.Width lido na primeira leitura de rotulos novos: "Sul 18" caixa 28,36 x rotulo 28,08; "Sudeste 30" 46,70 x 46,24;
' "Nordeste 9" 46,97 x 46,52 (a caixa fica ~0,5 pt acima).
Private Function MeasuredLabelWidth(cht As Chart, txt As String) As Double
    Dim shp As Shape
    Set shp = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 70, 16)
    With shp.TextFrame2
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeShapeToFitText
        .MarginLeft = 3: .MarginRight = 3: .MarginTop = 1: .MarginBottom = 1
        .TextRange.Text = txt
        .TextRange.Font.Size = 9
    End With
    MeasuredLabelWidth = shp.Width
    shp.Delete
End Function

' O rotulo do ultimo ponto de cada linha (a direita dele, ~5,5 pt depois) cabe antes da borda do grafico, com a largura MEDIDA do
' texto (MeasuredLabelWidth) limitada a 20% do grafico (o Excel quebra em linhas o rotulo mais largo: 455 pt, "Regiao Sudeste
' 1.400.000,00" em 2 linhas; 295 pt, em 3). A geometria do proprio DataLabel nao e lida (no Excel invisivel dos testes ela nao e
' atualizada). Nao usa o estimador nem as constantes de producao.
Private Sub AssertLineEndLabelFits(cht As Chart, tag As String)
    Dim s As Series, pt As Point, t As String, w As Double
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            Set pt = s.Points(s.Points.Count)
            t = pt.DataLabel.Text
            w = MeasuredLabelWidth(cht, t)
            If w > 0.2 * cht.ChartArea.Width Then w = 0.2 * cht.ChartArea.Width
            AssertTrue pt.Left + 5.5 + w <= cht.ChartArea.Width + 0.5, tag & ": rotulo '" & t & "' cabe (ponto=" & pt.Left & ", largura medida=" & w & ", grafico=" & cht.ChartArea.Width & ")"
        End If
    Next
End Sub

Private Function InsideRight(cht As Chart) As Double
    InsideRight = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth
End Function

' (I1) Tirar a legenda devolve o espaco mesmo com a area de plotagem em layout manual (fixada pela folga dos totais, das setas ou
' do fim da linha): o layout volta ao automatico e as folgas que so encolhem sao refeitas por cima. Posicoes reais da area interna.
Public Sub Test_Legend_None_GivesRoomBack_WhenPlotLayoutIsManual()
    Dim wb As Workbook, d As clsChartData, cht As Chart, w0 As Double, tot As Series
    ' barras 100% com totais: a folga dos totais fixou a area (layout manual)
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked100").Chart
    AssertEqual xlChartElementPositionCustom, cht.PlotArea.Position, "premissa: layout manual"
    w0 = cht.PlotArea.InsideWidth
    Deco_SetLegend cht, "none"
    cht.Refresh
    AssertTrue InsideRight(cht) >= cht.ChartArea.Width - 40, "sem legenda a area interna chega perto da borda, restando so a folga dos totais (direita=" & InsideRight(cht) & ", grafico=" & cht.ChartArea.Width & ")"
    AssertTrue cht.PlotArea.InsideWidth >= w0 + 40, "ganhou largura (antes " & w0 & ", agora " & cht.PlotArea.InsideWidth & ")"
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertTrue tot.Points(1).Left + 5.5 + MeasuredNumberWidth(tot.Points(1).DataLabel.Text) <= cht.ChartArea.Width, "o rotulo do total continua cabendo"
    CloseTemp wb
    ' area com seta de CAGR e totais: a folga de cima continua, a da direita volta
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Area_Insert(d.Source.Worksheet, d, "area").Chart
    Ann_AddCagr cht, 1, 4
    AssertEqual xlChartElementPositionCustom, cht.PlotArea.Position, "premissa: layout manual"
    Deco_SetLegend cht, "none"
    cht.Refresh
    AssertTrue InsideRight(cht) >= cht.ChartArea.Width - 15, "area: sem legenda a area interna chega a borda (direita=" & InsideRight(cht) & ")"
    AssertTrue cht.PlotArea.InsideTop >= 51.5, "a folga de cima da seta de CAGR com totais continua (topo interno=" & cht.PlotArea.InsideTop & ")"
    CloseTemp wb
    ' linha com o fim da linha: sair dele (none) devolve a largura que o rotulo tinha tomado
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "line").Chart
    Deco_SetLegend cht, "lineend"
    AssertTrue InsideRight(cht) < cht.ChartArea.Width - 15, "premissa: o rotulo do fim da linha tomou largura"
    Deco_SetLegend cht, "none"
    cht.Refresh
    AssertTrue InsideRight(cht) >= cht.ChartArea.Width - 15, "linha: de lineend para none a area interna volta a borda (direita=" & InsideRight(cht) & ")"
    Deco_SetLegend cht, "lineend"
    Deco_SetLegend cht, "right"
    cht.Refresh
    AssertTrue cht.PlotArea.Left + cht.PlotArea.Width <= cht.Legend.Left - 9, "de lineend para right a area de plotagem fica antes da legenda"
    AssertTrue InsideRight(cht) >= cht.Legend.Left - 40, "e nao sobra um vazio grande (direita=" & InsideRight(cht) & ", legenda=" & cht.Legend.Left & ")"
    CloseTemp wb
End Sub

' (I1) A mesma opcao de legenda duas vezes da o mesmo resultado (largura da area de plotagem estavel).
Public Sub Test_Legend_SameOptionTwice_KeepsPlotWidthStable()
    Dim wb As Workbook, d As clsChartData, cht As Chart, w1 As Double, kind As Variant
    For Each kind In Array("area", "line")
        Set wb = NewTempWorkbook(): Set d = Fixture(wb)
        If kind = "area" Then Set cht = Area_Insert(d.Source.Worksheet, d, "area").Chart Else Set cht = Stacked_Insert(d.Source.Worksheet, d, "line").Chart
        If kind = "area" Then Ann_AddCagr cht, 1, 4
        Deco_SetLegend cht, IIf(kind = "area", "right", "lineend")
        cht.Refresh
        w1 = cht.PlotArea.InsideWidth
        Deco_SetLegend cht, IIf(kind = "area", "right", "lineend")
        cht.Refresh
        AssertAlmost w1, cht.PlotArea.InsideWidth, 0.5, kind & ": largura estavel na segunda chamada"
        Deco_SetLegend cht, IIf(kind = "area", "right", "lineend")
        cht.Refresh
        AssertAlmost w1, cht.PlotArea.InsideWidth, 0.5, kind & ": e na terceira"
        CloseTemp wb
    Next
End Sub

' (I1) Barras com setas de CAGR e de diferenca: depois de tirar e religar a legenda a folga das setas continua (ponta a 10 pt do fim
' da maior barra, rotulo a direita da ponta e antes da borda livre: a do grafico sem legenda, a da legenda com ela).
Public Sub Test_Legend_Bars_Arrows_RoomKeptAcrossLegendChanges()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ln As Shape, lbl As Shape, ax As Axis, xEnd As Double, mode As Variant, edge As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Stacked_SetTotals cht, False
    Set ln = Ann_AddCagr(cht, 1, 4)
    Set lbl = cht.Shapes("tqq_ann_cagr_1_lbl")
    Ann_AddDifference cht, 1, 4
    For Each mode In Array("none", "right", "none", "right")
        Deco_SetLegend cht, CStr(mode)
    Core_SafeRefresh cht
        cht.Refresh
        Set ax = cht.Axes(xlValue)
        xEnd = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (57 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)      ' 2025: 18 + 30 + 9
        If mode = "right" Then edge = cht.Legend.Left Else edge = cht.ChartArea.Width
        AssertAlmost xEnd + 10, ln.Left + ln.Width, 0.8, mode & ": ponta da seta a 10 pt do fim da barra"
        AssertTrue lbl.Left >= ln.Left + ln.Width + 2, mode & ": rotulo a direita da ponta"
        AssertTrue lbl.Left + lbl.Width <= edge + 0.5, mode & ": rotulo antes da borda livre (direita=" & (lbl.Left + lbl.Width) & ", borda=" & edge & ")"
        AssertTrue cht.Shapes("tqq_ann_diff_1").Top + cht.Shapes("tqq_ann_diff_1").Height <= cht.ChartArea.Height, mode & ": seta de diferenca dentro do grafico"
    Next
    CloseTemp wb
End Sub

' (I2) Categorias que sao DATAS: o eixo e de texto (fatias iguais, mesmo com datas irregulares), o eixo "entre categorias" e fixado
' (Verdadeiro) e Geom_ChartBetween le esse valor; as auxiliares alinham pela regra de tres independente, a seta de diferenca fica a 10 pt
' do ponto da ultima categoria (area) ou da borda da coluna (combinacao).
Public Sub Test_Area_And_Combo_DateCategories_TextAxisAndAlignment()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, kind As Variant, ln As Shape, xTo As Double, i As Long, slot As Double, ln2 As Series, tot As Series
    Dim v As Variant, sm As Variant
    For Each kind In Array("area", "area100")
        Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
        ws.Range("A2").Value = DateSerial(2024, 1, 1): ws.Range("A3").Value = DateSerial(2024, 1, 2)
        ws.Range("A4").Value = DateSerial(2024, 3, 1): ws.Range("A5").Value = DateSerial(2024, 12, 1)
        ws.Range("A2:A5").NumberFormat = "dd/mm/yyyy"
        Set cht = Area_Insert(ws, d, CStr(kind)).Chart
        AssertEqual xlCategoryScale, cht.Axes(xlCategory).CategoryType, kind & ": eixo de texto (nao de tempo)"
        AssertTrue cht.Axes(xlCategory).AxisBetweenCategories, kind & ": eixo entre categorias fixado"
        AssertTrue Geom_ChartBetween(cht), kind & ": a geometria le o valor fixado"
        AssertAreaHelpersAligned cht, ws, (kind = "area100"), kind & " datas"
        If kind = "area" Then
            Set ln = Ann_AddDifference(cht, 1, 4)
            With cht.PlotArea
                xTo = .InsideLeft + .InsideWidth * 3.5 / 4
            End With
            AssertAlmost xTo + 10, ln.Left, 0.8, "area com datas: seta de diferenca a 10 pt do ponto da ultima categoria"
        End If
        CloseTemp wb
    Next
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A2").Value = DateSerial(2024, 1, 1): ws.Range("A3").Value = DateSerial(2024, 1, 2): ws.Range("A4").Value = DateSerial(2024, 12, 1)
    ws.Range("A2:A4").NumberFormat = "dd/mm/yyyy"
    Set cht = Combo_Insert(ws, d).Chart
    AssertEqual xlCategoryScale, cht.Axes(xlCategory).CategoryType, "combo: eixo de texto"
    AssertTrue cht.Axes(xlCategory).AxisBetweenCategories, "combo: eixo entre categorias fixado"
    AssertTrue Geom_ChartBetween(cht)
    Set ln2 = Combo_LineSeries(cht): Set tot = Chart_FindSeries(cht, "tqq_total")
    v = Array(0.42, 0.45, 0.47): sm = Array(160, 190, 230)
    For i = 1 To 3
        AssertAlmost ExpX(cht, i, 3), ln2.Points(i).Left, 0.8, "combo com datas: x da linha " & i
        AssertAlmost ExpX(cht, i, 3), tot.Points(i).Left, 0.8, "combo com datas: x do total " & i
        AssertAlmost ExpY(cht, CDbl(sm(i - 1)), xlPrimary), tot.Points(i).Top, 0.8, "combo com datas: y do total " & i
    Next
    Set ln = Ann_AddDifference(cht, 1, 3)
    slot = cht.PlotArea.InsideWidth / 3
    AssertAlmost cht.PlotArea.InsideLeft + slot * 2.5 + slot / 1.5 / 2 + 10, ln.Left, 0.8, "combo com datas: seta de diferenca a 10 pt da borda da ultima coluna"
    CloseTemp wb
End Sub

' (I3) Rollback da area: falha logo apos alocar o bloco E com o grafico ja criado (area e area100): nenhum grafico, nenhum nome
' tqq_ e uma insercao posterior funciona.
Public Sub Test_Area_Insert_RollsBack_BeforeAndAfterChartExists_AndLaterInsertWorks()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, kind As Variant, hook As Variant, en As Long, co As ChartObject
    For Each kind In Array("area", "area100")
        For Each hook In Array("block", "chart")
            Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
            If hook = "block" Then Util_FailAfterBlockForTest = True Else Util_FailAfterChartForTest = True
            On Error Resume Next
            Area_Insert ws, d, CStr(kind)
            en = Err.Number
            On Error GoTo 0
            Util_FailAfterBlockForTest = False: Util_FailAfterChartForTest = False
            AssertEqual vbObjectError + IIf(hook = "block", 9999, 9998), en, kind & "/" & hook & ": erro simulado propagado"
            AssertEqual 0, ws.ChartObjects.Count, kind & "/" & hook & ": nenhum grafico"
            AssertEqual 0, CountTqqNames(wb), kind & "/" & hook & ": nenhum nome tqq_"
            Set co = Area_Insert(ws, d, CStr(kind))
            AssertEqual 1, ws.ChartObjects.Count, kind & "/" & hook & ": insercao posterior funciona"
            AssertEqual 1, CountTqqNames(wb)
            AssertTrue Not Chart_FindSeries(co.Chart, "tqq_total") Is Nothing
            CloseTemp wb
        Next
    Next
End Sub

Private Function CountTqqNames(wb As Workbook) As Long
    Dim nm As Name
    For Each nm In wb.Names
        If InStr(nm.Name, ID_PREFIX) > 0 Then CountTqqNames = CountTqqNames + 1
    Next
End Function

' (I3) O Fail: do proprio Combo_Insert: o grafico ja existe (Stacked_Insert terminou) e o passo seguinte falha.
Public Sub Test_Combo_Insert_RollsBack_WhenFailureComesAfterChartExists()
    Dim wb As Workbook, d As clsChartData, en As Long, ws As Worksheet, co As ChartObject
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb): Set ws = d.Source.Worksheet
    Util_FailAfterChartForTest = True
    On Error Resume Next
    Combo_Insert ws, d
    en = Err.Number
    On Error GoTo 0
    Util_FailAfterChartForTest = False
    AssertEqual vbObjectError + 9998, en
    AssertEqual 0, ws.ChartObjects.Count, "grafico desfeito pelo Fail: do Combo_Insert"
    AssertEqual 0, CountTqqNames(wb), "bloco desfeito"
    Set co = Combo_Insert(ws, d)
    AssertEqual 1, ws.ChartObjects.Count, "insercao posterior funciona"
    AssertEqual 1, CountTqqNames(wb)
    AssertEqual xlSecondary, Combo_LineSeries(co.Chart).AxisGroup
    CloseTemp wb
End Sub

' (a) O rotulo "nome valor" do fim da linha ocupa no maximo ~20% do grafico (o Excel o quebra em linhas), entao a folga reservada tem esse teto
' (sem ele, um nome comprido tomaria ~35% da largura); e a area de plotagem nunca fica menor que 30% da largura do grafico nem <= 0.
Public Sub Test_Legend_LineEnd_RoomCappedAtWrapWidth_AndPlotKeepsMinimum()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, co As ChartObject
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Regiao Sudeste e Centro-Oeste do pais", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 100, 50): ws.Range("A3:C3").Value = Array(2023, 120, 60): ws.Range("A4:C4").Value = Array(2024, 140, 80)
    Set d = ChartData_FromRange(ws.Range("A1:C4"))
    Set co = Stacked_Insert(ws, d, "line"): Set cht = co.Chart
    Deco_SetLegend cht, "lineend"
    cht.Refresh
    AssertTrue cht.PlotArea.InsideWidth >= 380, "reserva limitada a ~20% do grafico (largura interna=" & cht.PlotArea.InsideWidth & ", sem o teto ficaria ~290)"
    AssertLineEndLabelFits cht, "455 pt"
    co.Width = 120
    Chart_EnsureRoomAfterChange cht
    cht.Refresh
    AssertTrue cht.PlotArea.Width >= 0.3 * cht.ChartArea.Width - 0.5, "grafico de 120 pt: a area de plotagem nao passa do minimo (largura=" & cht.PlotArea.Width & ", grafico=" & cht.ChartArea.Width & ")"
    AssertTrue cht.PlotArea.InsideWidth > 0, "largura interna positiva"
    CloseTemp wb
End Sub

' (f) "Alterar dados" para um grafico mais estreito com legenda a direita: a area de plotagem de tamanho manual (folga dos totais) nao
' fica sob a legenda (Chart_EnsureRoomAfterChange, depois do redimensionamento, inclui o ajuste a legenda).
Public Sub Test_Core_RebuildFromRange_NarrowerChart_PlotStaysLeftOfLegend()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, co As ChartObject
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set co = Stacked_Insert(ws, d, "barstacked100")
    co.Width = 250
    Set cht = co.Chart
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:D5"))
    Set cht = Core_ActiveChart()
    cht.Refresh
    AssertAlmost 250, cht.Parent.Width, 0.5, "premissa: o grafico novo tem a largura do antigo"
    AssertTrue cht.HasLegend
    AssertTrue cht.PlotArea.Left + cht.PlotArea.Width <= cht.Legend.Left - 9, "area de plotagem antes da legenda (direita=" & (cht.PlotArea.Left + cht.PlotArea.Width) & ", legenda=" & cht.Legend.Left & ")"
    Watchers_Stop
    CloseTemp wb
End Sub

' (g) A espessura da borda das areas e GRAVADA com ponto decimal em qualquer localidade (Str$ nunca usa a virgula; em pt-BR um
' CStr(1.5) daria "1,5") e lida com ponto ou virgula: o Mekko grava 1,5 pt e a troca de paleta a preserva.
Public Sub Test_Style_AreaBorder_WrittenWithDot_ReadWithEitherSeparator()
    Dim wb As Workbook, d As clsChartData, cht As Chart, en As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Area_Insert(d.Source.Worksheet, d, "area").Chart
    Style_SetAreaBorder cht, 1.5
    AssertEqual "1.5", Registry_GetOption(cht, "border"), "gravada com ponto"
    AssertEqual 1.5, cht.SeriesCollection(2).Format.Line.Weight, "aplicada as areas existentes"
    Style_ApplySeriesColors cht, "cinza"
    AssertEqual 1.5, cht.SeriesCollection(3).Format.Line.Weight, "a troca de paleta preserva"
    Registry_SetOption cht, "border", "2,5"                       ' escrito a mao com virgula: tambem e lido
    Style_ApplySeriesColors cht, "padrao"
    AssertEqual 2.5, cht.SeriesCollection(1).Format.Line.Weight, "valor com virgula"
    Style_SetAreaBorder cht, 0.75
    AssertEqual "0.75", Registry_GetOption(cht, "border")
    On Error Resume Next
    Style_SetAreaBorder cht, 0
    en = Err.Number
    On Error GoTo 0
    AssertEqual 5, en, "espessura nao positiva e recusada"
    AssertEqual "0.75", Registry_GetOption(cht, "border"), "recusa nao grava"
    CloseTemp wb
End Sub

' (h) Area com segmento NEGATIVO: o Excel empilha algebricamente (medido pelos pixels do grafico exportado: com 10, -8 e 5 os acumulados
' sao 10, 2 e 7; a faixa de -8 e o poligono entre 10 e 2 e a de 5, pintada depois, cobre 2..7; o que se ve acima de 7 e so a faixa de
' -8 ate 10). O topo VISIVEL e o maior acumulado (10), nao o Total (7) nem a soma dos positivos (15): TotalPos e a ancora das setas
' (Ann_GetCategoryData) seguem o topo visivel; o texto do rotulo do total continua sendo o Total (7); o Mid_j e o meio algebrico.
Public Sub Test_Area_NegativeSegment_AlgebraicStacking_TopLabelsAndCagrAnchor()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, blk As Range, body As Range, ln As Shape, v() As Double, t() As Double, ex As Double
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Ano", "Sul", "Sudeste", "Nordeste")
    ws.Range("A2:D2").Value = Array(2022, 10, -8, 5)
    ws.Range("A3:D3").Value = Array(2023, 12, 4, 6)
    ws.Range("A4:D4").Value = Array(2024, 15, 6, 8)
    ws.Range("A5:D5").Value = Array(2025, 18, 7, 9)
    Set d = ChartData_FromRange(ws.Range("A1:D5"))
    Set cht = Area_Insert(ws, d, "area").Chart
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 7, body.Cells(2, 2).Value, "Total = soma algebrica"
    AssertEqual 10, body.Cells(2, 3).Value, "TotalPos = maior acumulado (10, 2, 7)"
    AssertEqual 22, body.Cells(3, 3).Value, "sem negativo o TotalPos e o proprio Total"
    AssertEqual 5, body.Cells(2, Area_MidCol(blk, 1)).Value: AssertEqual 6, body.Cells(2, Area_MidCol(blk, 2)).Value, "meio algebrico de -8 entre 10 e 2"
    AssertAlmost 4.5, body.Cells(2, Area_MidCol(blk, 3)).Value
    AssertEqual "-8", Chart_FindSeries(cht, "tqq_lbl_2").Points(1).DataLabel.Text, "o rotulo da faixa mostra o valor negativo"
    AssertEqual "7", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "o texto do total e o Total"
    AssertAlmost ExpY(cht, 10), Chart_FindSeries(cht, "tqq_total").Points(1).Top, 0.8, "o rotulo do total ancora no topo visivel (10)"
    AssertAlmost ExpY(cht, 5), Chart_FindSeries(cht, "tqq_lbl_1").Points(1).Top, 0.8
    AssertAlmost ExpY(cht, 6), Chart_FindSeries(cht, "tqq_lbl_2").Points(1).Top, 0.8
    AssertAlmost ExpY(cht, 4.5), Chart_FindSeries(cht, "tqq_lbl_3").Points(1).Top, 0.8
    Ann_GetCategoryData cht, v, t
    AssertAlmost 7, v(1), 0.001, "valor usado nos textos = Total": AssertAlmost 10, t(1), 0.001, "topo da seta = maior acumulado"
    AssertAlmost 34, v(4), 0.001: AssertAlmost 34, t(4), 0.001
    Set ln = Ann_AddCagr(cht, 1, 4)
    ex = (34 / 7) ^ (1 / 3) - 1
    AssertEqual "CAGR " & Format$(ex, "+0.0%;-0.0%;0.0%"), cht.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text, "CAGR do Total (7 -> 34)"
    AssertAlmost ExpY(cht, 10) - 28, ln.Top + ln.Height, 0.8, "origem da seta a 10 pt + 18 pt (folga dos totais) acima do topo visivel (10), nao dos positivos (15)"
    AssertAlmost ExpY(cht, 34) - 28, ln.Top, 0.8, "ponta acima do topo da ultima categoria"
    CloseTemp wb
End Sub

' (h) #N/D numa celula de origem de area: o Excel plota 0 naquele ponto (a faixa some, nao interpola; medido nos pixels) e o bloco usa Val = 0:
' o ponto auxiliar do rotulo fica no topo da faixa anterior, o texto some (formato de zero vazio) e o total ignora o erro.
Public Sub Test_Area_NaSourceCell_ZeroThickness_HelperAtPreviousTop()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("C3").Formula = "=NA()"                                      ' Sudeste em 2023
    Set cht = Area_Insert(ws, d, "area").Chart
    AssertEqual "", Chart_FindSeries(cht, "tqq_lbl_2").Points(2).DataLabel.Text, "sem texto"
    AssertAlmost ExpY(cht, 12), Chart_FindSeries(cht, "tqq_lbl_2").Points(2).Top, 0.8, "auxiliar no topo da faixa anterior (Sul = 12)"
    AssertEqual "18", Chart_FindSeries(cht, "tqq_total").Points(2).DataLabel.Text, "total ignora o erro (12 + 6)"
    AssertAlmost ExpY(cht, 18), Chart_FindSeries(cht, "tqq_total").Points(2).Top, 0.8
    AssertAlmost ExpY(cht, 15), Chart_FindSeries(cht, "tqq_lbl_3").Points(2).Top, 0.8, "a faixa seguinte (6) parte do topo da anterior: meio em 12 + 3"
    CloseTemp wb
End Sub

' (d) Eixo secundario da combinacao: minimo em 0 quando todos os valores da linha sao >= 0 (o automatico escolheria 0,39 para 0,42..0,47
' e exageraria a inclinacao num eixo oculto), automatico com algum negativo; acompanha mudancas nos dados (Core_SafeRefresh).
Public Sub Test_Combo_SecondaryScale_StartsAtZeroUnlessNegative()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, ax As Axis
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb): Set ws = d.Source.Worksheet
    Set cht = Combo_Insert(ws, d).Chart
    Set ax = cht.Axes(xlValue, xlSecondary)
    AssertEqual 0, ax.MinimumScale: AssertFalse ax.MinimumScaleIsAuto, "minimo fixado em 0"
    AssertAlmost ExpY(cht, 0.45, xlSecondary), Combo_LineSeries(cht).Points(2).Top, 0.8, "a linha segue a escala 0..max"
    AssertTrue ax.MaximumScale >= 0.47, "maximo cobre o maior valor"
    ws.Range("D3").Value = -0.1                                          ' um valor negativo
    Core_SafeRefresh cht
    cht.Refresh
    Set ax = cht.Axes(xlValue, xlSecondary)
    AssertTrue ax.MinimumScaleIsAuto, "com negativo volta ao automatico"
    AssertTrue ax.MinimumScale <= -0.1, "e cobre o negativo (minimo=" & ax.MinimumScale & ")"
    ws.Range("D3").Value = 0.45
    Core_SafeRefresh cht
    cht.Refresh
    Set ax = cht.Axes(xlValue, xlSecondary)
    AssertEqual 0, ax.MinimumScale: AssertFalse ax.MinimumScaleIsAuto, "voltou a 0 com os dados positivos"
    Combo_SetSecondary cht, False
    Combo_SetSecondary cht, True
    AssertEqual 0, cht.Axes(xlValue, xlSecondary).MinimumScale, "religar o eixo secundario tambem fixa o minimo"
    CloseTemp wb
End Sub

' (c) Rotulos da linha da combinacao com a celula de origem em GERAL (que mostraria 0,423456789): regra de reserva sobre os valores da
' linha (inteiros = "int"; todos com |v| <= 1 = pct1; demais = dec2); com formato na celula (percentual) o rotulo segue a celula; e o
' formato acompanha mudancas feitas na celula depois (Core_SafeRefresh).
Public Sub Test_Combo_LineLabelFormat_GeneralCellFallbackRule_AndFollowsCellAfterwards()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, ln As Series, dec As String
    dec = Application.International(xlDecimalSeparator)
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb): Set ws = d.Source.Worksheet
    ws.Range("D2:D4").Value = Application.Transpose(Array(0.423456789, 0.45, 0.47))
    Set cht = Combo_Insert(ws, d).Chart
    Set ln = Combo_LineSeries(cht)
    AssertEqual "42" & dec & "3%", ln.Points(1).DataLabel.Text, "Geral com |v| <= 1: pct1 (nao 0,423456789)"
    ws.Range("D2:D4").NumberFormat = "0.0%"
    Core_SafeRefresh cht
    AssertEqual "42" & dec & "3%", ln.Points(1).DataLabel.Text, "celula em 0,0%: segue a celula"
    ws.Range("D2:D4").NumberFormat = "0%"
    Core_SafeRefresh cht
    AssertEqual "42%", ln.Points(1).DataLabel.Text, "mudou o formato da celula: o rotulo acompanha"
    ws.Range("D2:D4").NumberFormat = "General"
    Core_SafeRefresh cht
    AssertEqual "42" & dec & "3%", ln.Points(1).DataLabel.Text, "de volta a Geral: reserva"
    ws.Range("D2:D4").Value = Application.Transpose(Array(1234.5678, 2000.5, 99.1))
    Core_SafeRefresh cht
    AssertEqual "1" & Application.International(xlThousandsSeparator) & "234" & dec & "57", ln.Points(1).DataLabel.Text, "Geral com valores maiores que 1: dec2"
    ws.Range("D2:D4").Value = Application.Transpose(Array(100, 120, 150))
    Core_SafeRefresh cht
    AssertEqual "100", ln.Points(1).DataLabel.Text, "Geral com inteiros: int"
    CloseTemp wb
End Sub

' (f) A area de plotagem de tamanho manual (aqui fixada pela folga da seta de CAGR) acompanha o redimensionamento do grafico em escala,
' mas a legenda nao: o grafico mais estreito deixa a area SOB a legenda; Chart_EnsureRoomAfterChange (que "Alterar dados" chama depois do
' redimensionamento) inclui o ajuste a legenda. (No "Alterar dados" de barras a folga dos totais ja cobre; as setas nao sao preservadas.)
Public Sub Test_Chart_EnsureRoomAfterChange_FitsManualPlotToLegend_AfterResize()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, co As ChartObject
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Ano", "Regiao Sudeste e Centro-Oeste", "Regiao Norte e Nordeste", "Regiao Sul do pais")
    ws.Range("A2:D2").Value = Array(2022, 10, 20, 5): ws.Range("A3:D3").Value = Array(2023, 12, 22, 6)
    ws.Range("A4:D4").Value = Array(2024, 15, 25, 8): ws.Range("A5:D5").Value = Array(2025, 18, 30, 9)
    Set d = ChartData_FromRange(ws.Range("A1:D5"))
    Set co = Area_Insert(ws, d, "area"): Set cht = co.Chart
    Ann_AddCagr cht, 1, 4
    AssertEqual xlChartElementPositionCustom, cht.PlotArea.Position, "premissa: layout manual"
    co.Width = 250
    cht.Refresh
    AssertTrue cht.PlotArea.Left + cht.PlotArea.Width > cht.Legend.Left - 9.5, "premissa: o grafico mais estreito deixou a area de plotagem sob a legenda (direita=" & (cht.PlotArea.Left + cht.PlotArea.Width) & ", legenda=" & cht.Legend.Left & ")"
    Chart_EnsureRoomAfterChange cht
    cht.Refresh
    AssertTrue cht.PlotArea.Left + cht.PlotArea.Width <= cht.Legend.Left - 9.5, "depois do ajuste a area de plotagem fica antes da legenda (direita=" & (cht.PlotArea.Left + cht.PlotArea.Width) & ", legenda=" & cht.Legend.Left & ")"
    Chart_EnsureRoomAfterChange cht
    cht.Refresh
    AssertTrue cht.PlotArea.Left + cht.PlotArea.Width <= cht.Legend.Left - 9.5, "idempotente"
    CloseTemp wb
End Sub

' O minimo da area de plotagem (rede de seguranca): encolher alem de 30% da largura do grafico para no minimo e diz que nao ha mais o que reduzir.
Public Sub Test_Deco_ShrinkPlotWidth_StopsAtMinimumFraction()
    Dim wb As Workbook, d As clsChartData, cht As Chart, w0 As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Area_Insert(d.Source.Worksheet, d, "area").Chart
    w0 = cht.PlotArea.Width
    AssertTrue Deco_ShrinkPlotWidth(cht, 10), "encolhe 10 pt"
    AssertAlmost w0 - 10, cht.PlotArea.Width, 0.5
    AssertTrue Deco_ShrinkPlotWidth(cht, 5000), "um pedido enorme encolhe ate o minimo"
    AssertAlmost 0.3 * cht.ChartArea.Width, cht.PlotArea.Width, 0.5, "minimo = 30% da largura do grafico"
    AssertFalse Deco_ShrinkPlotWidth(cht, 20), "ja no minimo: nada a reduzir"
    AssertAlmost 0.3 * cht.ChartArea.Width, cht.PlotArea.Width, 0.5, "e a largura nao muda"
    AssertTrue cht.PlotArea.InsideWidth > 0
    CloseTemp wb
End Sub

' Tarefa 15b: um minimo definido a mao pelo usuario no eixo secundario da combinacao (fixo e diferente do 0 que o ThinQQ fixa) nao e
' sobrescrito por Combo_SyncSecondaryScale; voltando o minimo ao automatico, o ThinQQ volta a fixa-lo em 0. O 0 do proprio ThinQQ continua
' voltando ao automatico com um valor negativo.
Public Sub Test_Combo_SecondaryScale_KeepsUserManualMinimum()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, ax As Axis
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb): Set ws = d.Source.Worksheet
    Set cht = Combo_Insert(ws, d).Chart
    Set ax = cht.Axes(xlValue, xlSecondary)
    AssertEqual 0, ax.MinimumScale: AssertFalse ax.MinimumScaleIsAuto, "pre-condicao: o ThinQQ fixou 0"
    ax.MinimumScale = -0.5                                               ' o usuario fixa o minimo a mao
    Core_SafeRefresh cht
    Combo_SyncSecondaryScale cht
    Set ax = cht.Axes(xlValue, xlSecondary)
    AssertFalse ax.MinimumScaleIsAuto, "minimo manual continua fixo"
    AssertEqual -0.5, ax.MinimumScale, "minimo manual do usuario nao e sobrescrito"
    ws.Range("D3").Value = -0.1                                          ' dados com negativo: o minimo do usuario continua
    Core_SafeRefresh cht
    AssertEqual -0.5, cht.Axes(xlValue, xlSecondary).MinimumScale, "com negativo o minimo manual tambem fica"
    ws.Range("D3").Value = 0.45
    cht.Axes(xlValue, xlSecondary).MinimumScaleIsAuto = True             ' o usuario devolve ao automatico
    Core_SafeRefresh cht
    Set ax = cht.Axes(xlValue, xlSecondary)
    AssertEqual 0, ax.MinimumScale: AssertFalse ax.MinimumScaleIsAuto, "de volta ao automatico: o ThinQQ fixa 0 de novo"
    ws.Range("D3").Value = -0.1
    Core_SafeRefresh cht
    AssertTrue cht.Axes(xlValue, xlSecondary).MinimumScaleIsAuto, "o 0 do ThinQQ volta ao automatico com negativo"
    CloseTemp wb
End Sub

' Onda final (I-3): a coluna da linha da combinacao e localizada pela formula VIVA da serie (3o argumento de SERIES), nao pelo texto do
' endereco da origem gravado no bloco: renomear a planilha de origem nao faz os rotulos da linha (celula em Geral) virarem 0,42 a cada calculo.
Public Sub Test_Combo_LineLabels_KeepFallbackFormat_AfterSourceSheetRenamed()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, ln As Series, want As String
    Set wb = NewTempWorkbook(): Set d = ComboFixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:D4").Select
    Core_Insert "combo"
    Set cht = Core_ActiveChart()
    Set ln = Combo_LineSeries(cht)
    want = Format$(0.42, "0.0%")
    AssertEqual want, ln.Points(1).DataLabel.Text, "pre-condicao: Geral com razao vira pct1"
    ws.Name = "Origem renomeada"
    Core_SafeRefresh cht
    ws.Calculate
    Core_SafeRefresh cht
    AssertEqual want, ln.Points(1).DataLabel.Text, "depois de renomear a planilha de origem o formato de reserva continua"
    AssertEqual Format$(0.47, "0.0%"), ln.Points(3).DataLabel.Text
    ws.Range("D2:D4").NumberFormat = "0%"                            ' a celula ganha formato: o rotulo passa a segui-la
    Core_SafeRefresh cht
    AssertEqual Format$(0.42, "0%"), ln.Points(1).DataLabel.Text, "a celula com formato (planilha renomeada) e lida pela formula da serie"
    Watchers_Stop
    CloseTemp wb
End Sub

' Onda A, item E: tabela da demonstracao (12 meses de B2C + B2B com a margem em % na linha). O padrao (totais e eixo secundario) punha a linha
' sobre os rotulos dos totais. Agora a faixa da linha fica acima do MAIOR rotulo de total, conferida pela regra de tres nos eixos reais: o y de
' cada ponto da linha (escala secundaria) mais o raio do marcador fica acima do topo medido (ForceLayout) de todos os rotulos de total. Dois
' refresh nao mexem na escala (sem laco que cresce); colunas maiores continuam sob a linha.
Private Function DemoComboFixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet, i As Long
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Mês", "B2C", "B2B", "Margem")
    For i = 1 To 12
        ws.Cells(i + 1, 1).Value = DateSerial(2025, i, 1)
        ws.Cells(i + 1, 2).Value = 20000 + i * 700 + (i Mod 3) * 900
        ws.Cells(i + 1, 3).Value = 9000 + i * 350
        ws.Cells(i + 1, 4).Value = 0.3 + (i Mod 4) * 0.008
    Next
    ws.Range("A2:A13").NumberFormat = "mmm/yy"
    ws.Range("D2:D13").NumberFormat = "0.0%"
    Set DemoComboFixture = ChartData_FromRange(ws.Range("A1:D13"))
End Function

Private Sub AssertLineAboveTotals(cht As Chart, tag As String)
    Dim ln As Series, tot As Series, i As Long, v As Variant, topLbl As Double, y As Double
    ForceLayout cht
    Set ln = Combo_LineSeries(cht): Set tot = Chart_FindSeries(cht, "tqq_total")
    topLbl = 1E+30
    For i = 1 To tot.Points.Count
        If tot.Points(i).HasDataLabel Then
            If tot.Points(i).DataLabel.Top < topLbl Then topLbl = tot.Points(i).DataLabel.Top
        End If
    Next
    v = ln.Values
    For i = 1 To ln.Points.Count
        y = ExpY(cht, CDbl(v(LBound(v) + i - 1)), xlSecondary)
        AssertTrue y + 2.5 <= topLbl, tag & ": ponto " & i & " da linha (y=" & Format$(y, "0.0") & ") acima do maior rotulo de total (topo " & Format$(topLbl, "0.0") & ")"
        AssertAlmost y, ln.Points(i).Top + ln.Points(i).Height / 2, 1.5, tag & ": marcador " & i & " na escala secundaria real"
        AssertTrue ln.Points(i).DataLabel.Top >= -0.5, tag & ": rotulo da linha " & i & " dentro do grafico"
    Next
End Sub

Public Sub Test_Combo_LineAboveTotals_DemoMonths()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, sx As Axis, px As Axis, a1 As Double, a2 As Double, a3 As Double, i As Long
    Set wb = NewTempWorkbook(): Set d = DemoComboFixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:D13").Select
    Core_Insert "combo"
    Set cht = Core_ActiveChart()
    AssertEqual "1", Registry_GetOption(cht, "secondary"), "padrao: eixo secundario": AssertEqual "1", Registry_GetOption(cht, "totals"), "padrao: totais"
    AssertLineAboveTotals cht, "insercao"
    Set sx = cht.Axes(xlValue, xlSecondary): Set px = cht.Axes(xlValue, xlPrimary)
    a1 = sx.MinimumScale: a2 = sx.MaximumScale: a3 = px.MaximumScale
    Core_SafeRefresh cht
    Core_SafeRefresh cht
    AssertAlmost a1, sx.MinimumScale, 0.0000001, "estavel: minimo secundario": AssertAlmost a2, sx.MaximumScale, 0.0000001, "estavel: maximo secundario"
    AssertAlmost a3, px.MaximumScale, 0.0001, "estavel: maximo primario"
    For i = 2 To 13
        ws.Cells(i, 2).Value = ws.Cells(i, 2).Value * 1.5
    Next
    ws.Range("D5").Value = 0.36
    Core_SafeRefresh cht
    AssertLineAboveTotals cht, "colunas maiores e margem nova"
    AssertTrue px.MaximumScale > a3, "o maximo primario acompanha as colunas"
    Core_SetToggle "totals", False
    AssertTrue px.MaximumScaleIsAuto, "sem totais: escala primaria do Excel"
    AssertTrue sx.MaximumScaleIsAuto, "sem totais: maximo secundario do Excel"
    AssertEqual 0, sx.MinimumScale, "sem totais: o minimo 0 da Fase 2"
    Core_SetToggle "totals", True
    AssertLineAboveTotals cht, "totais de volta"
    sx.MaximumScale = 1                                         ' o usuario fixa o maximo secundario
    Core_SafeRefresh cht
    AssertEqual 1, sx.MaximumScale, "maximo secundario do usuario fica"
    AssertTrue px.MaximumScaleIsAuto, "e a folga das colunas sai"
    Watchers_Stop
    CloseTemp wb
End Sub

' Onda A, item F: com 12 trimestres de 5 digitos o rotulo da ponta (empurrado para dentro da faixa) encostava no vizinho ("3.5713.418"). Agora
' o rotulo da ponta que sobrepoe o vizinho da mesma faixa fica oculto (registrado em thinlbl) e volta quando cabe (grafico mais largo). Conferido
' pela geometria MEDIDA dos rotulos (ForceLayout): nenhum rotulo de ponta visivel sobrepoe o vizinho. Um rotulo apagado pelo usuario continua
' apagado; formato e paleta nao recriam os ocultos; religar que falha no meio deixa o rotulo oculto e na lista.
Private Function EndLabelsFixture(wb As Workbook) As Range
    Dim ws As Worksheet, i As Long
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Trimestre", "Cimento", "Tintas", "Hidráulica")
    For i = 1 To 12
        ws.Cells(i + 1, 1).Value = "T" & i
        ws.Cells(i + 1, 2).Value = 9200 + i * 37: ws.Cells(i + 1, 3).Value = 3571 - i * 13: ws.Cells(i + 1, 4).Value = 2400 + i * 11
    Next
    Set EndLabelsFixture = ws.Range("A1:D13")
End Function

Private Sub AssertEndLabelsClear(cht As Chart, tag As String)
    Dim j As Long, s As Series, e As Long, i As Long, nb As Long, a As DataLabel, b As DataLabel, list As String, sepH As Boolean, sepV As Boolean
    ForceLayout cht
    list = Registry_GetOption(cht, "thinlbl", "")
    For j = 1 To 3
        Set s = Chart_FindSeries(cht, "tqq_lbl_" & j)
        For e = 1 To 2
            If e = 1 Then i = 1: nb = 2 Else i = 12: nb = 11
            If s.Points(i).HasDataLabel And s.Points(nb).HasDataLabel Then
                Set a = s.Points(i).DataLabel: Set b = s.Points(nb).DataLabel
                sepH = (a.Left + a.Width <= b.Left + 0.5) Or (b.Left + b.Width <= a.Left + 0.5)
                sepV = (a.Top + a.Height <= b.Top + 0.5) Or (b.Top + b.Height <= a.Top + 0.5)
                AssertTrue sepH Or sepV, tag & ": rotulo da ponta " & j & "." & i & " sobre o vizinho (" & a.Left & "+" & a.Width & " / " & b.Left & ")"
            End If
        Next
    Next
End Sub

Public Sub Test_Area_EndLabels_HiddenWhenOverlapping_BackWhenFit()
    Dim wb As Workbook, rng As Range, cht As Chart, s1 As Series, s2 As Series, j As Long, en As Long, es As String, ed As String
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set rng = EndLabelsFixture(wb)
    rng.Select
    Core_Insert "area"
    Set cht = Core_ActiveChart()
    AssertTrue Len(Registry_GetOption(cht, "thinlbl", "")) > 0, "12 trimestres de 5 digitos: alguma ponta ocultada (" & Registry_GetOption(cht, "thinlbl", "") & ")"
    AssertEndLabelsClear cht, "estreito"
    For j = 1 To 3
        If Not Chart_FindSeries(cht, "tqq_lbl_" & j).Points(1).HasDataLabel Then AssertTrue InStr("," & Registry_GetOption(cht, "thinlbl", "") & ",", "," & j & ".1,") > 0, "oculto = registrado"
    Next
    Set s1 = Chart_FindSeries(cht, "tqq_lbl_1"): Set s2 = Chart_FindSeries(cht, "tqq_lbl_2")
    Core_SetOption "numfmt", "dec1"
    Core_SetOption "palette", "cinza"
    AssertEndLabelsClear cht, "formato e paleta"
    AssertTrue Len(Registry_GetOption(cht, "thinlbl", "")) > 0, "formato e paleta nao recriam os ocultos"
    Core_SetOption "numfmt", "int"
    cht.Parent.Width = 1400
    Core_SafeRefresh cht
    AssertEqual "", Registry_GetOption(cht, "thinlbl", ""), "largo: todos de volta"
    AssertTrue s1.Points(1).HasDataLabel And s1.Points(12).HasDataLabel
    AssertEqual xlLabelPositionRight, s1.Points(1).DataLabel.Position, "de volta para dentro da faixa"
    AssertEqual xlLabelPositionLeft, s1.Points(12).DataLabel.Position
    AssertEqual Format$(9237, "#,##0"), s1.Points(1).DataLabel.Text
    AssertEndLabelsClear cht, "largo"
    s2.Points(12).HasDataLabel = False                          ' o usuario apaga um rotulo de ponta
    cht.Parent.Width = 460
    Core_SafeRefresh cht
    cht.Parent.Width = 1400
    Core_SafeRefresh cht
    AssertFalse s2.Points(12).HasDataLabel, "apagado pelo usuario continua apagado"
    cht.Parent.Width = 460
    Core_SafeRefresh cht
    AssertFalse s1.Points(1).HasDataLabel, "estreito de novo: a ponta 1.1 oculta"
    Application.EnableEvents = False                            ' o refresh so quando o teste chama
    Util_FailThinSyncForTest = "area-link:1.1"
    cht.Parent.Width = 1400                                     ' (o redimensionamento pode disparar o refresh por evento: o gancho vem antes)
    Core_SafeRefresh cht
    AssertEqual "", Util_FailThinSyncForTest, "o gancho disparou"
    AssertFalse s1.Points(1).HasDataLabel, "religar falhou: oculto de novo"
    AssertTrue InStr("," & Registry_GetOption(cht, "thinlbl", "") & ",", ",1.1,") > 0, "e continua na lista"
    Core_SafeRefresh cht
    AssertTrue s1.Points(1).HasDataLabel, "a passada seguinte religa"
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailThinSyncForTest = ""
    Application.EnableEvents = True
    Watchers_Stop
    CloseTemp wb
    If en <> 0 Then Err.Raise en, es, ed
End Sub
