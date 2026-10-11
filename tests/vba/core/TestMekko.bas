Attribute VB_Name = "TestMekko"
Option Explicit

Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Regiao", "Eletro", "Moveis")
    ws.Range("A2:C2").Value = Array("Sul", 40, 10)
    ws.Range("A3:C3").Value = Array("Sudeste", 20, 10)
    ws.Range("A4:C4").Value = Array("Norte", 10, 10)
    Set Fixture = ChartData_FromRange(ws.Range("A1:C4"))
End Function

Private Function UnitsFixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Regiao", "Largura", "Eletro", "Moveis")
    ws.Range("A2:D2").Value = Array("Sul", 50, 40, 10)
    ws.Range("A3:D3").Value = Array("Sudeste", 30, 20, 10)
    ws.Range("A4:D4").Value = Array("Norte", 20, 10, 10)
    Set UnitsFixture = ChartData_FromRange(ws.Range("A1:D4"))
End Function

Private Function C(i As Long) As Long     ' linha do corpo (1-based, cabecalho = 1) do centro da categoria i
    C = 1 + Mekko_CenterRow(i)
End Function

Public Sub Test_Mekko_Block_XSharesAndHelpers()
    Dim wb As Workbook, blk As Range, body As Range
    Set wb = NewTempWorkbook()
    Set blk = Mekko_BuildBlock(wb, "mk0001", Fixture(wb), "mekko"): Set body = Block_Body(blk)
    AssertEqual 14, blk.Columns.Count: AssertEqual 16, body.Rows.Count, "corpo: cabecalho + 5 linhas por categoria (o bloco tem ainda a linha de meta)"
    AssertEqual 2, Mekko_SeriesCount(blk)
    ' X: larguras 50%, 30%, 20% -> bordas 1, 501, 801, 1001; centros 251, 651, 901
    AssertEqual 1, body.Cells(2, MK_COL_X).Value: AssertEqual 1, body.Cells(3, MK_COL_X).Value
    AssertEqual 251, body.Cells(C(1), MK_COL_X).Value
    AssertEqual 501, body.Cells(5, MK_COL_X).Value: AssertEqual 501, body.Cells(6, MK_COL_X).Value
    AssertEqual 501, body.Cells(7, MK_COL_X).Value: AssertEqual 651, body.Cells(C(2), MK_COL_X).Value
    AssertEqual 901, body.Cells(C(3), MK_COL_X).Value: AssertEqual 1001, body.Cells(16, MK_COL_X).Value
    AssertEqual "Sul", body.Cells(C(1), MK_COL_CAT).Value: AssertEqual "", CStr(body.Cells(2, MK_COL_CAT).Value)
    AssertEqual 50, body.Cells(C(1), MK_COL_W).Value: AssertEqual 50, body.Cells(C(1), MK_COL_TOTAL).Value
    ' Y: participacoes nas linhas L/C/R, zero em L0/R0
    AssertAlmost 0.8, body.Cells(C(1), Mekko_YCol(1)).Value: AssertAlmost 0.2, body.Cells(C(1), Mekko_YCol(2)).Value
    AssertEqual 0, body.Cells(2, Mekko_YCol(1)).Value: AssertAlmost 0.8, body.Cells(3, Mekko_YCol(1)).Value
    AssertEqual 0, body.Cells(6, Mekko_YCol(1)).Value, "R0 da categoria 1"
    AssertAlmost 2 / 3, body.Cells(C(2), Mekko_YCol(1)).Value, 0.000001
    ' auxiliares so nas linhas C
    AssertAlmost 0.4, body.Cells(C(1), Mekko_LblCol(blk, 1)).Value: AssertAlmost 0.9, body.Cells(C(1), Mekko_LblCol(blk, 2)).Value
    AssertTrue IsError(body.Cells(2, Mekko_LblCol(blk, 1)).Value)
    AssertAlmost 0.8, body.Cells(C(1), Mekko_LblValCol(blk, 1)).Value
    AssertEqual 1, body.Cells(C(1), Mekko_TotPosCol(blk)).Value: AssertAlmost 0.5, body.Cells(C(1), Mekko_TotLblCol(blk)).Value, 0.000001, "largura 50%"
    AssertEqual 0, body.Cells(C(1), Mekko_CatPosCol(blk)).Value: AssertTrue IsError(body.Cells(5, Mekko_CatPosCol(blk)).Value), "linha R: NA"
    AssertTrue IsError(body.Cells(2, Mekko_VLineCol(blk)).Value), "percent: sem linha de valor"
    CloseTemp wb
End Sub

Public Sub Test_Mekko_Insert_AreaOnDateAxis_LabelsTotalsCategories()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Mekko_Insert(d.Source.Worksheet, d, "mekko").Chart
    AssertEqual "mekko", Registry_ChartType(cht)
    AssertEqual xlAreaStacked, cht.SeriesCollection(1).ChartType
    AssertEqual xlTimeScale, cht.Axes(xlCategory).CategoryType
    AssertEqual 1, cht.Axes(xlCategory).MinimumScale: AssertEqual 1001, cht.Axes(xlCategory).MaximumScale
    AssertEqual 0, cht.Axes(xlValue).MinimumScale: AssertEqual 1, cht.Axes(xlValue).MaximumScale
    AssertEqual 15, cht.SeriesCollection(1).Points.Count
    AssertEqual "80%", Chart_FindSeries(cht, "tqq_lbl_1").Points(C(1) - 1).DataLabel.Text
    AssertEqual "20%", Chart_FindSeries(cht, "tqq_lbl_2").Points(C(1) - 1).DataLabel.Text
    AssertEqual "50%", Chart_FindSeries(cht, "tqq_total").Points(C(1) - 1).DataLabel.Text, "largura da coluna"
    AssertEqual "Sul", Chart_FindSeries(cht, "tqq_cat").Points(C(1) - 1).DataLabel.Text
    AssertEqual xlLabelPositionBelow, Chart_FindSeries(cht, "tqq_cat").Points(C(1) - 1).DataLabel.Position
    AssertEqual 2, cht.Legend.LegendEntries.Count
    AssertEqual "1", Registry_GetOption(cht, "ind100"): AssertEqual "3", Registry_GetOption(cht, "n")
    AssertEqual "pct0", Registry_GetOption(cht, "numfmt")
    AssertEqual RGB(255, 255, 255), cht.SeriesCollection(1).Format.Line.ForeColor.RGB, "separadores brancos"
    AssertEqual 1.5, cht.SeriesCollection(1).Format.Line.Weight, "separadores de 1,5 pt"
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    AssertEqual 1, CalcSheet_Cleanup(wb), "limpeza com o grafico novo (auxiliares com pontos NA) remove so o orfao"
    AssertFalse CalcSheet_CleanupReadFailed(), "pontos NA dos auxiliares nao sao falha de leitura"
    Mekko_SetLabels cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl_1") Is Nothing
    Mekko_SetTotals cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertTrue Not Chart_FindSeries(cht, "tqq_cat") Is Nothing, "categorias sempre"
    CloseTemp wb
End Sub

' Geometria independente: o centro de cada coluna e a soma das larguras anteriores + metade da propria.
Public Sub Test_Mekko_Geometry_Independent()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lbl As Series, il As Double, iw As Double, dl As DataLabel, centers As Variant, i As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Mekko_Insert(d.Source.Worksheet, d, "mekko").Chart
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth
    centers = Array(0.25, 0.65, 0.9)
    Set lbl = Chart_FindSeries(cht, "tqq_lbl_1")
    For i = 1 To 3
        Set dl = lbl.Points(C(i) - 1).DataLabel
        AssertAlmost il + iw * centers(i - 1), dl.Left + dl.Width / 2, 8, "rotulo da categoria " & i & " no centro da coluna"
    Next
    Set dl = Chart_FindSeries(cht, "tqq_total").Points(C(1) - 1).DataLabel
    AssertTrue dl.Top + dl.Height <= cht.PlotArea.InsideTop + 2, "total acima do nivel 100%"
    AssertTrue dl.Top >= 0, "rotulo de total dentro do grafico"
    AssertTrue cht.PlotArea.InsideTop >= 23.5, "folga acima da area de plotagem para os totais"
    AssertTrue Not cht.Shapes("tqq_ann_ind100") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_MekkoUnits_Block_VLine_AndNumFmt()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range, body As Range, s As Series, v As Variant
    Set wb = NewTempWorkbook(): Set d = UnitsFixture(wb)
    Set cht = Mekko_Insert(d.Source.Worksheet, d, "mekkounits").Chart
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual "mekkounits", Registry_ChartType(cht)
    AssertEqual 2, Mekko_SeriesCount(blk)
    AssertEqual 50, body.Cells(C(1), MK_COL_W).Value: AssertEqual 251, body.Cells(C(1), MK_COL_X).Value
    AssertEqual 40, body.Cells(C(1), Mekko_YCol(1)).Value, "unidades: valor absoluto"
    AssertEqual 50, body.Cells(C(1), Mekko_TotPosCol(blk)).Value: AssertEqual 50, body.Cells(C(1), Mekko_TotLblCol(blk)).Value
    AssertEqual 38, body.Cells(2, Mekko_VLineCol(blk)).Value, "media ponderada pela largura (2500+900+400)/100"
    AssertEqual "40", Chart_FindSeries(cht, "tqq_lbl_1").Points(C(1) - 1).DataLabel.Text
    AssertEqual "50", Chart_FindSeries(cht, "tqq_total").Points(C(1) - 1).DataLabel.Text
    AssertEqual "int", Registry_GetOption(cht, "numfmt"): AssertEqual "", Registry_GetOption(cht, "ind100", "")
    AssertTrue cht.Axes(xlValue).MaximumScaleIsAuto
    Deco_SetValueLine cht, True
    Set s = Chart_FindSeries(cht, "tqq_vline")
    v = s.Values
    AssertAlmost 38, CDbl(v(1)): AssertAlmost 38, CDbl(v(15))      ' CDbl: elemento de Variant passado ByRef a Double nao compila
    AssertEqual "Média 38", s.Points(15).DataLabel.Text
    CloseTemp wb
End Sub

' Review Focus 3: linha com total zero vira coluna de largura nula, sem #DIV/0! e sem coluna fantasma.
Public Sub Test_Mekko_ZeroTotalRow_NoErrorNoWidth()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range, r As Long, cc As Long, blk As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Worksheet.Range("B3:C3").Value = Array(0, 0)              ' Sudeste some: larguras 50/70 e 20/70
    Set cht = Mekko_Insert(d.Source.Worksheet, d, "mekko").Chart
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 715, body.Cells(5, MK_COL_X).Value, "borda direita de Sul = 1+ROUND(1000*50/70)"
    AssertEqual 715, body.Cells(C(2), MK_COL_X).Value, "categoria vazia tem largura zero"
    AssertEqual 715, body.Cells(10, MK_COL_X).Value: AssertEqual 1001, body.Cells(16, MK_COL_X).Value
    For r = 2 To 16
        For cc = 1 To 4 + Mekko_SeriesCount(blk)
            AssertFalse IsError(body.Cells(r, cc).Value), "erro em " & body.Cells(r, cc).Address
        Next
    Next
    AssertEqual 0, body.Cells(C(2), Mekko_YCol(1)).Value
    AssertEqual 15, cht.SeriesCollection(1).Points.Count
    CloseTemp wb
End Sub

Public Sub Test_Mekko_Refusals_Rollback_And_UnitsNeeds3Cols()
    Dim wb As Workbook, d As clsChartData, d2 As clsChartData, ws As Worksheet, cht As Chart, en As Long, ed As String, nm As Name
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Mekko_Insert(ws, d, "mekko").Chart
    On Error Resume Next
    Ann_AddDifference cht, 1, 3
    en = Err.Number: ed = Err.Description: Err.Clear
    AssertEqual vbObjectError + 7001, en: AssertEqual "A seta de diferença não está disponível para este tipo de gráfico.", ed
    Deco_SetValueLine cht, True
    en = Err.Number: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en, "linha de valor so em unidades"
    Set d2 = ChartData_FromRange(ws.Range("A1:B4"))          ' 2 colunas: largura sem nenhuma serie
    On Error Resume Next
    Mekko_Insert ws, d2, "mekkounits"
    en = Err.Number: ed = Err.Description: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 6009, en
    AssertEqual "Mekko (unidades): a segunda coluna da tabela é a largura; são necessárias pelo menos 3 colunas.", ed
    AssertEqual 1, ws.ChartObjects.Count, "nada inserido na recusa"
    ws.ChartObjects(1).Delete: CalcSheet_Cleanup wb
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Mekko_Insert ws, d, "mekko"
    en = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, en
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    CloseTemp wb
End Sub

Public Sub Test_Core_Insert_Mekko_TogglesAndOptions()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:C4").Select
    Core_Insert "mekko"
    Set cht = Core_ActiveChart()
    AssertEqual "mekko", Registry_ChartType(cht)
    AssertTrue Core_Supports("ind100"): AssertFalse Core_Supports("gap"): AssertFalse Core_Supports("cagr")
    Core_SetOption "numfmt", "pct1"
    AssertEqual "80" & Application.International(xlDecimalSeparator) & "0%", Chart_FindSeries(cht, "tqq_lbl_1").Points(C(1) - 1).DataLabel.Text
    Core_SetToggle "ind100", False
    AssertFalse Core_GetToggle("ind100")
    Core_SetToggle "totals", False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), cht.SeriesCollection(1).Format.Fill.ForeColor.RGB
    AssertEqual 1.5, cht.SeriesCollection(1).Format.Line.Weight, "trocar a paleta mantem os separadores do Mekko"
    Core_ClearLastNotice                                          ' o contador conta desde aqui, nao desde testes anteriores
    Core_Refresh
    AssertEqual 0, Core_NoticeCount(), "sem aviso de tabela mudada (n = 3 categorias, 15 pontos)"
    Watchers_Stop
    CloseTemp wb
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Alem do brief (carry-overs da Tarefa 9)
' ---------------------------------------------------------------------------------------------------------------------

' A geometria de um rotulo de dados lida depois de mudar o layout pode ser a antiga ate o grafico ser desenhado (Tarefa 8): exportar desenha.
Private Sub ForceLayout(cht As Chart)
    cht.Export TempDir() & "\layout.png"
End Sub

' So os pontos C (3, 8, 13, ...) das auxiliares tem rotulo: um ponto NA() com rotulo mostraria "0" no zero (medido).
Private Sub AssertNoStrayLabels(cht As Chart, tag As String)
    Dim nm As Variant, s As Series, r As Long
    For Each nm In Array("tqq_lbl_1", "tqq_lbl_2", "tqq_total", "tqq_cat")
        Set s = Chart_FindSeries(cht, CStr(nm))
        If Not s Is Nothing Then
            For r = 1 To s.Points.Count
                AssertEqual ((r - 3) Mod MK_ROWS_PER_CAT = 0), s.Points(r).HasDataLabel, tag & ": " & nm & " ponto " & r
            Next
        End If
    Next
End Sub

Private Function CountTqqNames(wb As Workbook) As Long
    Dim nm As Name
    For Each nm In wb.Names
        If Left$(nm.Name, Len(ID_PREFIX)) = ID_PREFIX Then CountTqqNames = CountTqqNames + 1
    Next
End Function

' Carry-overs (a) e (h): o indicador 100% fica no nivel 100% logo depois de inserir; a folga de cima (totais) e a de baixo (nomes) sobrevivem a
' legenda (layout automatico) e ao redimensionamento; formato e paleta escrevem por ponto (nao recriam rotulos nos pontos NA()); mudar um valor
' redistribui as larguras ao vivo (regra de tres sobre a area interna real); rollback quando a falha vem depois de o grafico existir.
Public Sub Test_Mekko_Room_Indicator_LiveWidths_AndRollbackAfterChart()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, dl As DataLabel, il As Double, iw As Double, i As Long
    Dim centers As Variant, en As Long, n0 As Long, nNames As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Mekko_Insert(ws, d, "mekko").Chart
    AssertAlmost cht.PlotArea.InsideTop, cht.Shapes("tqq_ann_ind100").Top, 1, "tique do 100% no topo da area interna logo apos inserir"
    AssertNoStrayLabels cht, "inicial"
    AssertEqual RGB(255, 255, 255), Chart_FindSeries(cht, "tqq_lbl_2").Points(C(1) - 1).DataLabel.Font.Color, "padrao: faixa 2 azul, fonte branca"
    cht.Parent.Select
    Core_SetOption "palette", "tema"                 ' faixa 2 = destaque 2 do tema (laranja claro no Office): a fonte vira cinza-escuro
    Core_SetOption "numfmt", "pct1"
    AssertNoStrayLabels cht, "depois de paleta e formato"
    AssertEqual Style_ContrastFont(wb.Theme.ThemeColorScheme.Colors(msoThemeAccent2).RGB), Chart_FindSeries(cht, "tqq_lbl_2").Points(C(1) - 1).DataLabel.Font.Color, "contraste com a cor nova da faixa 2"
    AssertEqual 1.5, cht.SeriesCollection(2).Format.Line.Weight, "separadores de 1,5 pt depois da paleta"
    AssertEqual "1.5", Registry_GetOption(cht, "border"), "carry-over (f): espessura gravada na opcao, com ponto em qualquer localidade"
    Core_SetToggle "totals", False
    Core_SetToggle "totals", True
    AssertEqual "50%", Chart_FindSeries(cht, "tqq_total").Points(C(1) - 1).DataLabel.Text, "totais religados pelo Core: largura da coluna"
    AssertNoStrayLabels cht, "depois de religar os totais"
    ForceLayout cht
    Set dl = Chart_FindSeries(cht, "tqq_cat").Points(C(1) - 1).DataLabel
    AssertTrue dl.Top + dl.Height <= cht.ChartArea.Height + 0.5, "nome da categoria dentro do grafico (fim=" & (dl.Top + dl.Height) & ")"
    AssertTrue dl.Top >= cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight + 2, "nome abaixo do eixo, nao preso na borda do grafico"
    Deco_SetLegend cht, "none"
    AssertTrue cht.PlotArea.InsideTop >= 23.5, "folga de cima sem legenda"
    Deco_SetLegend cht, "right"
    AssertTrue cht.PlotArea.InsideTop >= 23.5, "folga de cima com a legenda de volta"
    AssertTrue cht.PlotArea.Left + cht.PlotArea.Width <= cht.Legend.Left, "area de plotagem antes da legenda"
    cht.Parent.Width = 320: cht.Parent.Height = 210
    Chart_EnsureRoomAfterChange cht
    AssertTrue cht.PlotArea.InsideTop >= 23.5, "folga de cima no grafico menor"
    AssertTrue cht.ChartArea.Height - (cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight) >= 20.5, "folga de baixo para os nomes"
    ws.Range("B4").Value = 70                        ' totais 50, 30, 80 de 160
    cht.Refresh
    ForceLayout cht
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth
    centers = Array(25 / 160, (50 + 15) / 160, (80 + 40) / 160)
    For i = 1 To 3
        Set dl = Chart_FindSeries(cht, "tqq_cat").Points(C(i) - 1).DataLabel
        AssertAlmost il + iw * centers(i - 1), dl.Left + dl.Width / 2, 3, "nome da categoria " & i & " no centro da coluna depois de mudar os dados"
    Next
    AssertEqual "50%", Chart_FindSeries(cht, "tqq_total").Points(C(3) - 1).DataLabel.Text, "largura de Norte: 80/160"
    n0 = ws.ChartObjects.Count: nNames = CountTqqNames(wb)
    Util_FailAfterChartForTest = True
    On Error Resume Next
    Mekko_Insert ws, d, "mekko"
    en = Err.Number
    On Error GoTo 0
    Util_FailAfterChartForTest = False
    AssertEqual vbObjectError + 9998, en, "falha simulada depois do grafico"
    AssertEqual n0, ws.ChartObjects.Count, "o grafico novo foi apagado"
    AssertEqual nNames, CountTqqNames(wb), "o bloco novo foi apagado"
    Watchers_Stop
    CloseTemp wb
End Sub

' Unidades, geometria pela escala REAL do eixo (regra de tres, independente das formulas): totais acima do topo de cada coluna e no centro
' dela, rotulo do segmento no meio da faixa, rotulo da media depois da borda direita (o ultimo ponto e X = 1001, na borda: o eixo de datas
' nao tem meia fatia) e antes da legenda; carry-over (c): o topo por ponto (Mesma escala) e a altura da coluna; largura texto = coluna nula
' sem rotulos soltos (Review Focus 3).
Public Sub Test_MekkoUnits_Geometry_TotalsSegmentsVLine_AndZeroWidthNoGhost()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, dl As DataLabel, ax As Axis, body As Range
    Dim il As Double, iw As Double, it As Double, ih As Double, mn As Double, mx As Double, i As Long, centers As Variant, totals As Variant
    Dim v() As Double, t() As Double
    Set wb = NewTempWorkbook(): Set d = UnitsFixture(wb): Set ws = d.Source.Worksheet
    Set cht = Mekko_Insert(ws, d, "mekkounits").Chart
    Deco_SetValueLine cht, True
    ForceLayout cht
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth: it = cht.PlotArea.InsideTop: ih = cht.PlotArea.InsideHeight
    Set ax = cht.Axes(xlValue): mn = ax.MinimumScale: mx = ax.MaximumScale
    centers = Array(0.25, 0.65, 0.9): totals = Array(50, 30, 20)
    For i = 1 To 3
        Set dl = Chart_FindSeries(cht, "tqq_total").Points(C(i) - 1).DataLabel
        AssertAlmost il + iw * centers(i - 1), dl.Left + dl.Width / 2, 3, "total " & i & " no centro da coluna"
        AssertTrue dl.Top + dl.Height <= it + ih * (mx - totals(i - 1)) / (mx - mn) + 1, "total " & i & " acima do topo da coluna"
        AssertTrue dl.Top >= 0, "total " & i & " dentro do grafico"
    Next
    Set dl = Chart_FindSeries(cht, "tqq_lbl_1").Points(C(1) - 1).DataLabel
    AssertAlmost it + ih * (mx - 20) / (mx - mn), dl.Top + dl.Height / 2, 3, "Eletro de Sul (0..40): rotulo no meio da faixa"
    Set dl = Chart_FindSeries(cht, "tqq_vline").Points(15).DataLabel
    AssertTrue dl.Left >= il + iw - 1, "rotulo da media depois da borda direita da area interna"
    AssertTrue dl.Left + dl.Width <= cht.Legend.Left - 3, "rotulo da media antes da legenda (fim=" & (dl.Left + dl.Width) & ", legenda=" & cht.Legend.Left & ")"
    Ann_GetCategoryData cht, v, t
    AssertEqual 15, UBound(t), "um topo por ponto"
    AssertAlmost 50, t(C(1) - 1): AssertAlmost 0, t(1), , "L0 no zero"
    ws.Range("B3").Value = "abc"                     ' largura texto: Sudeste sem largura
    cht.Refresh
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual body.Cells(5, MK_COL_X).Value, body.Cells(C(2), MK_COL_X).Value, "Sudeste: centro na borda de Sul"
    AssertEqual "", Chart_FindSeries(cht, "tqq_cat").Points(C(2) - 1).DataLabel.Text, "sem nome solto"
    AssertEqual "", Chart_FindSeries(cht, "tqq_total").Points(C(2) - 1).DataLabel.Text, "sem total solto"
    AssertEqual "", Chart_FindSeries(cht, "tqq_lbl_1").Points(C(2) - 1).DataLabel.Text, "sem segmento solto"
    AssertEqual "Sul", Chart_FindSeries(cht, "tqq_cat").Points(C(1) - 1).DataLabel.Text
    AssertEqual "Média 41", Chart_FindSeries(cht, "tqq_vline").Points(15).DataLabel.Text, "media so das colunas com largura: (2500+400)/70"
    ' area empilhada: o Excel soma algebricamente (Tarefa 6); com -10 e depois 40 o topo visivel e 30, nao a soma dos positivos (40)
    ws.Range("C2:D2").Value = Array(-10, 40)
    cht.Refresh
    Ann_GetCategoryData cht, v, t
    AssertAlmost 30, t(C(1) - 1), , "topo visivel = maior acumulado"
    CloseTemp wb
End Sub

' Carry-over (g): "Alterar dados" reaplica as escolhas salvas no Mekko (formato, paleta e borda, totais, indicador, legenda) e nas unidades
' (linha de valor, rotulos), com a folga de cima refeita e o n da tabela nova.
Public Sub Test_Core_RebuildFromRange_Mekko_ReappliesSavedOptions()
    Dim wb As Workbook, d As clsChartData, du As clsChartData, ws As Worksheet, cht As Chart, nc As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A5:C5").Value = Array("Centro", 30, 30)
    Set cht = Mekko_Insert(ws, d, "mekko").Chart
    cht.Parent.Select
    Core_SetOption "numfmt", "pct1"
    Core_SetOption "palette", "cinza"
    Core_SetToggle "totals", False
    Core_SetToggle "ind100", False
    Core_SetOption "legend", "none"
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:C5")), "recriado com 4 categorias"
    Set nc = Core_ActiveChart()
    AssertEqual "mekko", Registry_ChartType(nc)
    AssertEqual "4", Registry_GetOption(nc, "n"): AssertEqual 20, nc.SeriesCollection(1).Points.Count
    AssertFalse Core_ChartSourceMismatch(nc)
    AssertEqual "80" & Application.International(xlDecimalSeparator) & "0%", Chart_FindSeries(nc, "tqq_lbl_1").Points(C(1) - 1).DataLabel.Text
    AssertEqual RGB(64, 64, 64), nc.SeriesCollection(1).Format.Fill.ForeColor.RGB
    AssertEqual 1.5, nc.SeriesCollection(1).Format.Line.Weight, "borda do Mekko copiada"
    AssertTrue Chart_FindSeries(nc, "tqq_total") Is Nothing, "totais continuam desligados"
    AssertFalse Ann_HasInd100(nc): AssertEqual "0", Registry_GetOption(nc, "ind100")
    AssertFalse nc.HasLegend: AssertEqual "none", Registry_GetOption(nc, "legend")
    AssertTrue nc.PlotArea.InsideTop >= 23.5, "folga de cima refeita"
    AssertTrue Not Chart_FindSeries(nc, "tqq_cat") Is Nothing
    AssertNoStrayLabels nc, "recriado"
    ws.Range("F1:I1").Value = Array("Regiao", "Largura", "Eletro", "Moveis")
    ws.Range("F2:I2").Value = Array("Sul", 50, 40, 10)
    ws.Range("F3:I3").Value = Array("Sudeste", 30, 20, 10)
    ws.Range("F4:I4").Value = Array("Norte", 20, 10, 10)
    ws.Range("F4").Value = DateSerial(2026, 3, 1): ws.Range("F4").NumberFormat = "mmm/yy"     ' categoria data: o nome segue o formato da origem
    Set du = ChartData_FromRange(ws.Range("F1:I4"))
    Set cht = Mekko_Insert(ws, du, "mekkounits").Chart
    Deco_SetValueLine cht, True
    Mekko_SetLabels cht, False
    AssertTrue Core_RebuildFromRange(cht, ws.Range("F1:I4"))
    Set nc = Core_ActiveChart()
    AssertEqual "mekkounits", Registry_ChartType(nc)
    AssertTrue Not Chart_FindSeries(nc, "tqq_vline") Is Nothing, "linha de valor religada"
    AssertEqual "Média 38", Chart_FindSeries(nc, "tqq_vline").Points(15).DataLabel.Text
    AssertTrue Chart_FindSeries(nc, "tqq_lbl_1") Is Nothing, "rotulos continuam desligados"
    AssertTrue Not Chart_FindSeries(nc, "tqq_total") Is Nothing
    AssertEqual "", Registry_GetOption(nc, "ind100", ""), "unidades: sem indicador 100%"
    AssertEqual ws.Range("F4").Text, Chart_FindSeries(nc, "tqq_cat").Points(C(3) - 1).DataLabel.Text, "categoria data com o formato da origem"
    ' tabela mudada: as series do Mekko apontam para o bloco (sempre 5n pontos); o aviso vem das referencias das categorias no bloco
    AssertFalse Core_ChartSourceMismatch(nc), "tabela do mesmo tamanho"
    ws.Rows(1).Insert                                ' acima das tabelas: dados nas linhas 3 a 5
    AssertFalse Core_ChartSourceMismatch(nc), "linha inserida acima da tabela nao muda o tamanho"
    ws.Rows(4).Insert                                ' entre a primeira e a segunda categoria
    AssertTrue Core_ChartSourceMismatch(nc), "linha inserida no meio da tabela"
    Watchers_Stop
    CloseTemp wb
End Sub
