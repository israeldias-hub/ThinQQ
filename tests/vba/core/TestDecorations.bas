Attribute VB_Name = "TestDecorations"
Option Explicit
' Decoracoes da Fase 2 (Tarefa 7): linha de valor na media, indicador 100% e mesma escala. Geometria sempre por regra de tres sobre o
' PlotArea/eixo reais (nunca pelas funcoes de producao) e pelos pontos das series auxiliares (Point.Left/Top, confiaveis; a geometria
' de um DataLabel nao e: ver a Tarefa 6).

Private Function Fixture(wb As Workbook, Optional mult As Double = 1) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Ano", "Sul", "Sudeste", "Nordeste")
    ws.Range("A2:D2").Value = Array(2022, 10 * mult, 20 * mult, 5 * mult)
    ws.Range("A3:D3").Value = Array(2023, 12 * mult, 22 * mult, 6 * mult)
    ws.Range("A4:D4").Value = Array(2024, 15 * mult, 25 * mult, 8 * mult)
    ws.Range("A5:D5").Value = Array(2025, 18 * mult, 30 * mult, 9 * mult)
    Set Fixture = ChartData_FromRange(ws.Range("A1:D5"))
End Function

Private Function CountPrefix(cht As Chart, p As String) As Long
    Dim i As Long
    For i = 1 To cht.Shapes.Count
        If Left$(cht.Shapes(i).Name, Len(p)) = p Then CountPrefix = CountPrefix + 1
    Next
End Function

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

' Total de cada categoria da tabela de origem (lido da PLANILHA, nunca do bloco).
Private Function RowTotal(ws As Worksheet, r As Long, lastCol As Long) As Double
    RowTotal = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(r, 2), ws.Cells(r, lastCol)))
End Function

Public Sub Test_VLine_Stacked_MeanLineWithLinkedLabel()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, v As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    Deco_SetValueLine cht, True
    Set s = Chart_FindSeries(cht, "tqq_vline")
    AssertTrue Not s Is Nothing
    v = s.Values
    AssertAlmost 45, CDbl(v(1)): AssertAlmost 45, CDbl(v(4))
    AssertEqual msoTrue, s.Format.Line.Visible
    AssertEqual msoLineDash, s.Format.Line.DashStyle
    AssertEqual "Média 45", s.Points(4).DataLabel.Text
    AssertFalse s.Points(1).HasDataLabel, "so o ultimo ponto tem rotulo"
    AssertEqual "1", Registry_GetOption(cht, "vline")
    AssertEqual 3, cht.Legend.LegendEntries.Count, "fora da legenda"
    d.Source.Worksheet.Range("B5").Value = 118          ' totais 35, 40, 48, 157 -> media 70
    v = s.Values
    AssertAlmost 70, CDbl(v(1)), 0.001, "a media e formula do bloco"
    AssertEqual "Média 70", s.Points(4).DataLabel.Text
    Deco_SetValueLine cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_vline") Is Nothing: AssertEqual "0", Registry_GetOption(cht, "vline")
    CloseTemp wb
End Sub

Public Sub Test_VLine_Bars_IsScatterAcrossCategories()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, x As Variant, y As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barclustered").Chart
    Deco_SetValueLine cht, True
    Set s = Chart_FindSeries(cht, "tqq_vline")
    AssertEqual xlSecondary, s.AxisGroup
    x = s.XValues: y = s.Values
    AssertAlmost 45, CDbl(x(1)), 0.001: AssertAlmost 45, CDbl(x(4)), 0.001
    AssertAlmost 3.5, CDbl(y(1)): AssertAlmost 0.5, CDbl(y(4))
    AssertEqual "Média 45", s.Points(4).DataLabel.Text
    AssertEqual msoTrue, s.Format.Line.Visible
    CloseTemp wb
End Sub

Public Sub Test_VLine_Refused_On100AndWaterfall()
    Dim wb As Workbook, d As clsChartData, cht As Chart, en As Long, ed As String, ws As Worksheet
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Stacked_Insert(ws, d, "stacked100").Chart
    On Error Resume Next
    Deco_SetValueLine cht, True
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    AssertEqual "A linha de valor não está disponível para este tipo de gráfico.", ed
    AssertTrue Chart_FindSeries(cht, "tqq_vline") Is Nothing
    Set cht = Waterfall_Insert(ws, d).Chart
    On Error Resume Next
    Deco_SetValueLine cht, True
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en, "cascata: sem linha de valor"
    AssertEqual "A linha de valor não está disponível para este tipo de gráfico.", ed
    AssertTrue Chart_FindSeries(cht, "tqq_vline") Is Nothing
    ws.Range("A1:D5").Select
    Core_Insert "stacked100"
    Core_ClearLastNotice
    Core_SetToggle "vline", True
    ' via API o aviso fala da linha de valor (a Tarefa 11 troca o texto pela mensagem generica da matriz,
    ' "'Linha de valor (media)' nao se aplica..."; a chamada direta acima mantem a mensagem especifica)
    AssertTrue InStr(Core_LastNotice(), "inha de valor") > 0, Core_LastNotice()
    AssertFalse Core_GetToggle("vline")
    Watchers_Stop
    CloseTemp wb
End Sub

Public Sub Test_Ind100_DefaultOn_Geometry_And_Toggle()
    Dim wb As Workbook, d As clsChartData, cht As Chart, tick As Shape, lbl As Shape, yTop As Double, ax As Axis
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked100").Chart
    AssertEqual "1", Registry_GetOption(cht, "ind100"), "ligado por padrao nos 100%"
    Set tick = cht.Shapes("tqq_ann_ind100"): Set lbl = cht.Shapes("tqq_ann_ind100_lbl")
    AssertEqual "100%", lbl.TextFrame2.TextRange.Text
    Set ax = cht.Axes(xlValue)
    yTop = cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight * (ax.MaximumScale - 1) / (ax.MaximumScale - ax.MinimumScale)
    AssertAlmost yTop, tick.Top, 1, "tique no nivel 100%"
    AssertTrue tick.Left + tick.Width <= cht.PlotArea.InsideLeft, "tique a esquerda da area de plotagem"
    AssertTrue lbl.Left + lbl.Width <= tick.Left + 0.5, "rotulo a esquerda do tique"
    AssertTrue lbl.Left >= 0, "dentro do grafico (folga criada)"
    Deco_SetInd100 cht, False
    AssertEqual 0, CountPrefix(cht, "tqq_ann_ind100"): AssertEqual "0", Registry_GetOption(cht, "ind100")
    Deco_SetInd100 cht, True
    AssertEqual 2, CountPrefix(cht, "tqq_ann_ind100")
    Ann_RemoveAll cht
    AssertEqual 2, CountPrefix(cht, "tqq_ann_ind100"), "Remover setas nao remove o indicador"
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked100").Chart
    Set tick = cht.Shapes("tqq_ann_ind100")
    AssertAlmost cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth, tick.Left, 1, "barras: tique no fim (100%) no topo"
    AssertTrue tick.Top + tick.Height <= cht.PlotArea.InsideTop + 0.5
    CloseTemp wb
End Sub

Public Sub Test_Ind100_Refused_OnStacked()
    Dim wb As Workbook, d As clsChartData, cht As Chart, en As Long, ed As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    AssertEqual "", Registry_GetOption(cht, "ind100", "")
    On Error Resume Next
    Deco_SetInd100 cht, True
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    AssertEqual "O indicador 100% só existe em gráficos 100% e Mekko (%).", ed
    AssertEqual 0, CountPrefix(cht, "tqq_ann_ind100")
    CloseTemp wb
End Sub

Public Sub Test_SameScale_Union_Persist_Refresh_Reset()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, c1 As Chart, c2 As Chart, coll As New Collection, skipped As Long, n As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set c1 = Stacked_Insert(ws, d, "stacked").Chart
    ws.Range("F1:I1").Value = Array("Ano", "A", "B", "C")
    ws.Range("F2:I2").Value = Array(2022, 100, 200, 50): ws.Range("F3:I3").Value = Array(2023, 120, 220, 60)
    Set c2 = Stacked_Insert(ws, ChartData_FromRange(ws.Range("F1:I3")), "clustered").Chart
    coll.Add c1: coll.Add c2
    n = Deco_SameScale(coll, skipped)
    AssertEqual 2, n: AssertEqual 0, skipped
    AssertEqual c2.Axes(xlValue).MaximumScale, c1.Axes(xlValue).MaximumScale
    AssertEqual c2.Axes(xlValue).MinimumScale, c1.Axes(xlValue).MinimumScale
    AssertTrue c1.Axes(xlValue).MaximumScale >= 220, "uniao cobre o maior grafico"
    AssertTrue Len(Registry_GetOption(c1, "vmax")) > 0 And Len(Registry_GetOption(c1, "vmin")) > 0
    AssertAlmost c1.Axes(xlValue).MaximumScale, Val(Registry_GetOption(c1, "vmax")), 0.000001, "texto neutro com ponto"
    c1.Axes(xlValue).MaximumScaleIsAuto = True                   ' alguem mexeu: o refresh reaplica
    Core_SafeRefresh c1
    AssertEqual c2.Axes(xlValue).MaximumScale, c1.Axes(xlValue).MaximumScale, "escala fixa reaplicada no refresh"
    Deco_ResetScale c1
    AssertTrue c1.Axes(xlValue).MaximumScaleIsAuto
    AssertEqual "", Registry_GetOption(c1, "vmax", "")
    CloseTemp wb
End Sub

Public Sub Test_Core_SameScale_SheetAndSelection_Notices()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, c3 As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:D5").Select: Core_Insert "stacked"
    ws.Range("A1:D5").Select: Core_Insert "clustered"
    ws.Range("A1:D5").Select: Core_Insert "stacked100"
    Set c3 = Core_ActiveChart()
    ws.Range("A1").Select
    Core_ClearLastNotice
    Core_SameScale
    AssertEqual "Mesma escala aplicada a 2 gráfico(s). 1 ignorado(s): tipo sem escala livre (100%, Mekko (%), pizza, Gantt).", Core_LastNotice()
    AssertEqual "", Registry_GetOption(c3, "vmax", ""), "100% ignorado"
    Deco_ResetScale ws.ChartObjects(1).Chart: Deco_ResetScale ws.ChartObjects(2).Chart
    ws.Shapes.Range(Array(ws.ChartObjects(1).Name, ws.ChartObjects(2).Name)).Select
    Core_ClearLastNotice
    Core_SameScale
    AssertEqual "Mesma escala aplicada a 2 gráfico(s).", Core_LastNotice()
    AssertEqual ws.ChartObjects(1).Chart.Axes(xlValue).MaximumScale, ws.ChartObjects(2).Chart.Axes(xlValue).MaximumScale
    ws.ChartObjects(1).Select
    Core_ClearLastNotice
    Core_ResetScale
    AssertEqual "Escala automática restaurada em 1 gráfico(s).", Core_LastNotice()
    AssertTrue ws.ChartObjects(1).Chart.Axes(xlValue).MaximumScaleIsAuto
    Teardown wb
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Linha de valor: geometria, familias, bloco antigo, formato, folga
' ---------------------------------------------------------------------------------------------------------------------

' A linha cai onde a regra de tres sobre o PlotArea e o eixo reais manda: colunas (y = media, x = centro de cada categoria) e barras
' (x = media no eixo primario, y = centro de cada faixa; categoria 1 no topo), e continua alinhada depois de mudar o MAXIMO do eixo
' primario a mao (a auxiliar XY segue o eixo; era a lacuna da Tarefa 3) e depois de mudar os dados.
Public Sub Test_VLine_Geometry_ColumnsAndBars_RuleOfThree_AfterManualMaximum()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, s As Series, tot As Series, ax As Axis, i As Long
    Dim il As Double, iw As Double, it As Double, ih As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Stacked_Insert(ws, d, "stacked").Chart
    Deco_SetValueLine cht, True
    Set s = Chart_FindSeries(cht, "tqq_vline"): Set ax = cht.Axes(xlValue)
    AssertEqual msoTrue, s.Points(4).DataLabel.Format.Fill.Visible, "rotulo com fundo branco (cai sobre a ultima coluna)"
    AssertEqual RGB(255, 255, 255), s.Points(4).DataLabel.Format.Fill.ForeColor.RGB
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth: it = cht.PlotArea.InsideTop: ih = cht.PlotArea.InsideHeight
    For i = 1 To 4
        AssertAlmost il + iw * (i - 0.5) / 4, s.Points(i).Left, 1, "colunas: x do ponto " & i
        AssertAlmost it + ih * (ax.MaximumScale - 45) / (ax.MaximumScale - ax.MinimumScale), s.Points(i).Top, 1, "colunas: y do ponto " & i
    Next
    ax.MaximumScale = 120                                        ' maximo manual: a linha continua na media
    Core_SafeRefresh cht
    it = cht.PlotArea.InsideTop: ih = cht.PlotArea.InsideHeight
    AssertAlmost it + ih * (120 - 45) / (120 - ax.MinimumScale), s.Points(1).Top, 1, "colunas: y depois do maximo manual 120"
    CloseTemp wb

    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Stacked_Insert(ws, d, "barstacked").Chart
    Deco_SetValueLine cht, True
    Set s = Chart_FindSeries(cht, "tqq_vline"): Set tot = Chart_FindSeries(cht, "tqq_total"): Set ax = cht.Axes(xlValue, xlPrimary)
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth: it = cht.PlotArea.InsideTop: ih = cht.PlotArea.InsideHeight
    For i = 1 To 4
        AssertAlmost il + iw * (45 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale), s.Points(i).Left, 1.5, "barras: x do ponto " & i
        AssertAlmost it + ih * (i - 0.5) / 4, s.Points(i).Top, 1.5, "barras: y do ponto " & i
    Next
    ax.MaximumScale = 120                                        ' maximo manual: tanto a linha quanto os totais (XY) seguem o eixo
    Core_SafeRefresh cht
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth
    For i = 1 To 4
        AssertAlmost il + iw * (45 - ax.MinimumScale) / (120 - ax.MinimumScale), s.Points(i).Left, 1.5, "barras: x da linha depois do maximo manual, ponto " & i
        AssertAlmost il + iw * (RowTotal(ws, i + 1, 4) - ax.MinimumScale) / (120 - ax.MinimumScale), tot.Points(i).Left, 1.5, "barras: x do total " & i & " depois do maximo manual"
    Next
    ws.Range("B5").Value = 118                                   ' total 157: a escala manual fica (120): a linha vai para a nova media (70)
    Core_SafeRefresh cht
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth
    AssertAlmost il + iw * (70 - ax.MinimumScale) / (120 - ax.MinimumScale), s.Points(1).Left, 1.5, "barras: a linha acompanha a media nova"
    CloseTemp wb
End Sub

' Area e combinacao: a linha usa a coluna VLine do bloco da familia (Area_VLineCol / Stacked_VLineCol) e a media e dos totais das
' COLUNAS (na combinacao a ultima coluna da tabela e a linha, fora dos totais).
Public Sub Test_VLine_AreaAndCombo_UseFamilyColumns()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, s As Series, v As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Area_Insert(ws, d, "area").Chart
    Deco_SetValueLine cht, True
    Set s = Chart_FindSeries(cht, "tqq_vline")
    AssertTrue Not s Is Nothing
    v = s.Values
    AssertAlmost 45, CDbl(v(1)), 0.001: AssertAlmost 45, CDbl(v(4)), 0.001
    AssertEqual "Média 45", s.Points(4).DataLabel.Text
    AssertFalse s.Points(2).HasDataLabel
    AssertEqual 3, cht.Legend.LegendEntries.Count, "area: fora da legenda"
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Combo_Insert(ws, d).Chart
    Deco_SetValueLine cht, True
    Set s = Chart_FindSeries(cht, "tqq_vline")
    v = s.Values
    AssertAlmost 38, CDbl(v(1)), 0.001, "combinacao: media dos totais das colunas (30, 34, 40, 48), sem a linha"
    AssertEqual "Média 38", s.Points(4).DataLabel.Text
    AssertEqual 3, cht.Legend.LegendEntries.Count, "combinacao: 2 colunas + a linha, fora a auxiliar"
    CloseTemp wb
End Sub

' Bloco da Fase 1 (nunca publicado: sem as colunas Pos/VLine): recusa clara, nada criado. Simulado trocando o cabecalho da coluna.
Public Sub Test_VLine_Phase1Block_Refused_NoSeries()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range, en As Long, ed As String, col As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    Set body = Block_Body(Registry_Block(cht))
    col = Stacked_VLineCol(Registry_Block(cht))
    AssertEqual "VLine", CStr(body.Cells(1, col).Value)
    AssertFalse Deco_VLineCells(cht) Is Nothing
    body.Cells(1, col).Value = "Part3"                           ' o que o bloco antigo teria na ultima coluna
    AssertTrue Deco_VLineCells(cht) Is Nothing, "cabecalho invalido: sem celulas"
    On Error Resume Next
    Deco_SetValueLine cht, True
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    AssertTrue InStr(ed, "versão anterior") > 0, ed
    AssertTrue Chart_FindSeries(cht, "tqq_vline") Is Nothing
    AssertEqual "", Registry_GetOption(cht, "vline", "")
    body.Cells(1, col).Value = "VLine"
                                                                 ' barras: a coluna Pos tambem e validada
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Set body = Block_Body(Registry_Block(cht))
    body.Cells(1, Stacked_PosCol(Registry_Block(cht))).Value = "x"
    On Error Resume Next
    Deco_SetValueLine cht, True
    en = Err.Number
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en, "barras sem a coluna Pos"
    AssertTrue Chart_FindSeries(cht, "tqq_vline") Is Nothing
    CloseTemp wb
End Sub

' O rotulo da media usa o formato do grafico (numfmt) e mostra "Média" tambem com media negativa; religar nao duplica a serie.
Public Sub Test_VLine_LabelFormat_FollowsNumfmt_NegativeMean_Idempotent()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, s As Series, dec As String, n As Long, sr As Series
    Dim neg As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    dec = Application.International(xlDecimalSeparator)
    ws.Range("A1:D5").Select: Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetToggle "vline", True
    Core_SetOption "numfmt", "dec1"
    Set s = Chart_FindSeries(cht, "tqq_vline")
    AssertEqual "Média 45" & dec & "0", s.Points(4).DataLabel.Text, "o formato do grafico vale para a media"
    Deco_SetValueLine cht, True                                   ' religar: refaz a serie, sem duplicar
    For Each sr In cht.SeriesCollection
        If sr.Name = "tqq_vline" Then n = n + 1
    Next
    AssertEqual 1, n, "uma so serie tqq_vline"
    Core_SetOption "numfmt", "int"
    ws.Range("B2:D5").Value = -10                                 ' totais -30: media -30
    Core_SafeRefresh cht
    Set s = Chart_FindSeries(cht, "tqq_vline")
    neg = s.Points(4).DataLabel.Text
    AssertEqual "Média -30", neg, "media negativa tambem mostra o texto"
    ws.Range("B2:D5").Value = 0                                   ' media 0: aparece (a secao de zero dos rotulos e vazia)
    Core_SafeRefresh cht
    AssertEqual "Média 0", s.Points(4).DataLabel.Text, "media zero aparece"
    Teardown wb
End Sub

' Folga do rotulo "Média N" (a direita do ultimo ponto, a meia fatia da borda): com muitas categorias e numeros grandes o rotulo
' nao pode passar da borda livre (a legenda), e a folga sobrevive a troca da legenda (que volta o layout ao automatico).
Public Sub Test_VLine_LabelRoom_BeforeLegend_AndAcrossLegendChanges()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, s As Series, i As Long, txt As String, w As Double
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "A", "B")
    For i = 1 To 8
        ws.Cells(i + 1, 1).Value = 2017 + i
        ws.Cells(i + 1, 2).Value = 1000000 * i
        ws.Cells(i + 1, 3).Value = 2000000 * i
    Next
    Set d = ChartData_FromRange(ws.Range("A1:C9"))
    Set cht = Stacked_Insert(ws, d, "stacked").Chart
    Deco_SetValueLine cht, True
    Set s = Chart_FindSeries(cht, "tqq_vline")
    txt = s.Points(8).DataLabel.Text
    AssertEqual "Média " & Format$(13500000, "#,##0"), txt         ' media dos totais 3.000.000 * (1..8) = 13.500.000 (separador da localidade)
    w = 6 + 5 * Len(txt)
    If w > 0.2 * cht.ChartArea.Width Then w = 0.2 * cht.ChartArea.Width     ' o Excel quebra o que passa de ~20% do grafico
    AssertTrue s.Points(8).Left + 6 + w <= cht.Legend.Left + 0.5, "rotulo antes da legenda: ponto=" & s.Points(8).Left & " largura~" & w & " legenda=" & cht.Legend.Left
    Deco_SetLegend cht, "none"
    AssertTrue s.Points(8).Left + 6 + w <= cht.ChartArea.Width + 0.5, "sem legenda: rotulo dentro do grafico"
    Deco_SetLegend cht, "right"
    AssertTrue s.Points(8).Left + 6 + w <= cht.Legend.Left + 0.5, "a folga sobrevive a legenda trocada: ponto=" & s.Points(8).Left & " legenda=" & cht.Legend.Left
    CloseTemp wb
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Indicador 100%
' ---------------------------------------------------------------------------------------------------------------------

' Area 100% (colunas): tique no nivel 100% a esquerda; as folgas sobrevivem a troca da legenda (que volta o plot ao automatico);
' o indicador convive com as setas (Ann_RemoveAll leva as setas, nao o indicador) e acompanha o refresh.
Public Sub Test_Ind100_Area100_RoomSurvivesLegend_CoexistsWithArrows()
    Dim wb As Workbook, d As clsChartData, cht As Chart, tick As Shape, lbl As Shape, ax As Axis, y As Double, mode As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Area_Insert(d.Source.Worksheet, d, "area100").Chart
    AssertEqual "1", Registry_GetOption(cht, "ind100"), "area 100%: ligado por padrao"
    Ann_AddCagr cht, 1, 4
    For Each mode In Array("none", "right", "none", "right")
        Deco_SetLegend cht, CStr(mode)
        Core_SafeRefresh cht                                     ' como Core_SetOption: o refresh reposiciona as formas
        Set tick = cht.Shapes("tqq_ann_ind100"): Set lbl = cht.Shapes("tqq_ann_ind100_lbl"): Set ax = cht.Axes(xlValue)
        y = cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight * (ax.MaximumScale - 1) / (ax.MaximumScale - ax.MinimumScale)
        AssertAlmost y, tick.Top, 1, "legenda " & mode & ": tique no nivel 100%"
        AssertTrue cht.PlotArea.InsideLeft >= 46 - 0.5, "legenda " & mode & ": folga a esquerda >= 46 (InsideLeft=" & cht.PlotArea.InsideLeft & ")"
        AssertTrue lbl.Left >= -0.5, "legenda " & mode & ": rotulo dentro do grafico (left=" & lbl.Left & ")"
        AssertTrue lbl.Left + lbl.Width <= tick.Left + 0.5, "legenda " & mode & ": rotulo a esquerda do tique"
    Next
    Ann_RemoveAll cht
    AssertEqual 0, CountPrefix(cht, "tqq_ann_cagr")
    AssertEqual 2, CountPrefix(cht, "tqq_ann_ind100")
    Ann_Refresh cht
    AssertAlmost y, cht.Shapes("tqq_ann_ind100").Top, 1, "o refresh mantem o tique"
    CloseTemp wb
End Sub

' Barras 100%: o rotulo "100%" fica centrado acima do tique (que sobe do fim da barra), dentro do grafico (folga de cima),
' e a folga de cima sobrevive a legenda trocada.
Public Sub Test_Ind100_Bars_LabelCenteredAboveTick_RoomSurvivesLegend()
    Dim wb As Workbook, d As clsChartData, cht As Chart, tick As Shape, lbl As Shape, x As Double, mode As Variant, ax As Axis
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "barstacked100").Chart
    For Each mode In Array("right", "none", "right")
        Deco_SetLegend cht, CStr(mode)
        Core_SafeRefresh cht                                     ' como Core_SetOption: o refresh reposiciona as formas
        Set tick = cht.Shapes("tqq_ann_ind100"): Set lbl = cht.Shapes("tqq_ann_ind100_lbl"): Set ax = cht.Axes(xlValue, xlPrimary)
        x = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (1 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
        AssertAlmost x, tick.Left, 1, "legenda " & mode & ": tique no fim da barra 100%"
        AssertAlmost x, lbl.Left + lbl.Width / 2, 1.5, "legenda " & mode & ": rotulo centrado no tique"
        AssertTrue lbl.Top >= -0.5, "legenda " & mode & ": rotulo dentro do grafico (top=" & lbl.Top & ")"
        AssertTrue lbl.Top + lbl.Height <= tick.Top + 0.5, "legenda " & mode & ": rotulo acima do tique"
        AssertTrue cht.PlotArea.InsideTop >= 20 - 0.5, "legenda " & mode & ": folga de cima >= 20"
    Next
    CloseTemp wb
End Sub

' O indicador so vale nos tipos com o recurso: se o tipo do bloco nao o tem, o refresh oculta as formas (valido = matriz).
Public Sub Test_Ind100_HiddenByRefresh_WhenKindHasNoIndicator()
    Dim wb As Workbook, d As clsChartData, cht As Chart, blk As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked100").Chart
    AssertEqual msoTrue, cht.Shapes("tqq_ann_ind100").Visible
    Set blk = Registry_Block(cht)
    blk.Cells(1, 2).Value = "stacked"                              ' simula um tipo sem o recurso
    Ann_Refresh cht
    AssertEqual msoFalse, cht.Shapes("tqq_ann_ind100").Visible, "tique oculto"
    AssertEqual msoFalse, cht.Shapes("tqq_ann_ind100_lbl").Visible, "rotulo oculto"
    blk.Cells(1, 2).Value = "stacked100"
    Ann_Refresh cht
    AssertEqual msoTrue, cht.Shapes("tqq_ann_ind100").Visible
    AssertEqual msoTrue, cht.Shapes("tqq_ann_ind100_lbl").Visible
    CloseTemp wb
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Mesma escala: bordas
' ---------------------------------------------------------------------------------------------------------------------

' Barras com mesma escala: as auxiliares XY (totais e linha de valor) continuam alinhadas com o fim das barras depois de a escala
' primaria mudar (o grafico menor ganha o maximo do maior), e uma segunda aplicacao da mesma escala e estavel.
Public Sub Test_SameScale_Bars_XYHelpersAlignedAfterScaleChange()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, c1 As Chart, c2 As Chart, coll As New Collection, skipped As Long
    Dim tot As Series, vl As Series, ax As Axis, i As Long, il As Double, iw As Double, mx As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set c1 = Stacked_Insert(ws, d, "barstacked").Chart
    Deco_SetValueLine c1, True
    ws.Range("F1:I1").Value = Array("Ano", "A", "B", "C")
    ws.Range("F2:I2").Value = Array(2022, 100, 200, 50): ws.Range("F3:I3").Value = Array(2023, 120, 220, 60)
    Set c2 = Stacked_Insert(ws, ChartData_FromRange(ws.Range("F1:I3")), "barstacked").Chart
    coll.Add c1: coll.Add c2
    AssertEqual 2, Deco_SameScale(coll, skipped)
    Set ax = c1.Axes(xlValue, xlPrimary)
    AssertTrue ax.MaximumScale >= 400, "o maximo do grafico maior (total 400)"
    AssertEqual c2.Axes(xlValue, xlPrimary).MaximumScale, ax.MaximumScale
    Set tot = Chart_FindSeries(c1, "tqq_total"): Set vl = Chart_FindSeries(c1, "tqq_vline")
    il = c1.PlotArea.InsideLeft: iw = c1.PlotArea.InsideWidth
    For i = 1 To 4
        AssertAlmost il + iw * (RowTotal(ws, i + 1, 4) - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale), tot.Points(i).Left, 1.5, "total " & i & " no fim da barra com a escala comum"
    Next
    AssertAlmost il + iw * (45 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale), vl.Points(1).Left, 1.5, "linha de valor com a escala comum"
    mx = ax.MaximumScale
    AssertEqual 2, Deco_SameScale(coll, skipped)                    ' de novo: mesma escala
    AssertEqual mx, c1.Axes(xlValue, xlPrimary).MaximumScale, "estavel ao repetir"
    CloseTemp wb
End Sub

' Escala fixa no maximo dos dados + seta de CAGR em barras: a barra mais longa termina no fim do eixo e a seta e o rotulo precisam da
' folga cheia a direita (80 + extra), que o Ann_EnsureRoom da criacao (maximo automatico) nao deu; aplicar a escala refaz a folga.
Public Sub Test_SameScale_FixedMaximum_CagrRoomRedone_Bars()
    Dim wb As Workbook, d As clsChartData, c1 As Chart, ln As Shape, lbl As Shape, ax As Axis, tip As Double, edge As Double, tot As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set c1 = Stacked_Insert(d.Source.Worksheet, d, "barstacked").Chart
    Ann_AddCagr c1, 1, 4
    Registry_SetOption c1, "vmin", "0"
    Registry_SetOption c1, "vmax", Deco_ScaleText(57)               ' o maior total: a barra 4 termina no fim do eixo
    Core_SafeRefresh c1                                            ' ApplyFixedScale: fixa o eixo e refaz as folgas
    Set ax = c1.Axes(xlValue, xlPrimary)
    AssertFalse ax.MaximumScaleIsAuto
    AssertAlmost 57, ax.MaximumScale, 0.000001
    Set ln = c1.Shapes("tqq_ann_cagr_1"): Set lbl = c1.Shapes("tqq_ann_cagr_1_lbl")
    edge = c1.Legend.Left
    tot = 57
    tip = c1.PlotArea.InsideLeft + c1.PlotArea.InsideWidth * (tot - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale) + 10 + 18       ' fim da barra + 10 + extra dos totais (>= 18)
    AssertTrue ln.Left + ln.Width >= tip - 1.5, "a ponta nao foi comprimida: ponta=" & (ln.Left + ln.Width) & " esperado>=" & tip
    AssertTrue lbl.Left + lbl.Width <= edge + 0.5, "o rotulo da seta cabe antes da legenda: " & (lbl.Left + lbl.Width) & " > " & edge
    AssertTrue lbl.Left >= ln.Left + ln.Width - 1, "o rotulo fica depois da ponta (nao foi empurrado sobre ela)"
    CloseTemp wb
End Sub

' So um grafico com escala livre (os demais sao 100%, pizza...): nada e feito e nada e gravado; a cascata tem escala livre.
Public Sub Test_SameScale_NeedsTwoFreeScales_And_WaterfallCounts()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, c1 As Chart, c2 As Chart, c3 As Chart, coll As Collection, skipped As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set c1 = Stacked_Insert(ws, d, "stacked").Chart
    Set c2 = Stacked_Insert(ws, d, "stacked100").Chart
    Set coll = New Collection: coll.Add c1: coll.Add c2
    AssertEqual 0, Deco_SameScale(coll, skipped), "um grafico com escala livre e um 100%: nada a igualar"
    AssertEqual 1, skipped
    AssertEqual "", Registry_GetOption(c1, "vmax", ""), "nada gravado"
    AssertTrue c1.Axes(xlValue).MaximumScaleIsAuto, "o eixo continua automatico"
    Set c3 = Waterfall_Insert(ws, d).Chart
    coll.Add c3
    AssertEqual 2, Deco_SameScale(coll, skipped), "coluna empilhada + cascata"
    AssertEqual 1, skipped
    AssertEqual c1.Axes(xlValue).MaximumScale, c3.Axes(xlValue).MaximumScale
    AssertTrue Len(Registry_GetOption(c3, "vmax", "")) > 0
    Registry_SetOption c2, "vmin", "0": Registry_SetOption c2, "vmax", "5"
    Deco_ApplyFixedScale c2                                        ' 100%: nunca fixa o eixo, mesmo com a opcao gravada
    AssertTrue c2.Axes(xlValue).MaximumScaleIsAuto, "100%: o eixo continua automatico"
    CloseTemp wb
End Sub

' A escala gravada nunca derruba o refresh: texto invalido, intervalo invertido, opcao vazia.
Public Sub Test_ApplyFixedScale_InvalidOptions_NeverThrow_AndLeaveAxisAlone()
    Dim wb As Workbook, d As clsChartData, cht As Chart, mx As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    mx = cht.Axes(xlValue).MaximumScale
    Deco_ApplyFixedScale cht                                       ' sem opcao
    Registry_SetOption cht, "vmin", "abc": Registry_SetOption cht, "vmax", "xyz"
    Deco_ApplyFixedScale cht                                       ' Val("abc") = 0 = Val("xyz"): intervalo vazio, ignorado
    Registry_SetOption cht, "vmin", "100": Registry_SetOption cht, "vmax", "10"
    Deco_ApplyFixedScale cht                                       ' invertido
    Registry_SetOption cht, "vmin", "": Registry_SetOption cht, "vmax", "80"
    Deco_ApplyFixedScale cht                                       ' meio vazio
    AssertTrue cht.Axes(xlValue).MaximumScaleIsAuto, "nenhuma escala invalida chegou ao eixo"
    AssertAlmost mx, cht.Axes(xlValue).MaximumScale, 0.000001
    Core_SafeRefresh cht
    CloseTemp wb
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' API do core
' ---------------------------------------------------------------------------------------------------------------------

Public Sub Test_Core_SelectedCharts_ActiveChart_And_MultiSelection()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, coll As Collection
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:D5").Select: Core_Insert "stacked"
    ws.Range("A1:D5").Select: Core_Insert "clustered"
    ws.Range("A1:D5").Select: Core_Insert "line"
    ws.Range("A1").Select
    Set coll = Core_SelectedCharts()
    AssertEqual 0, coll.Count, "sem grafico selecionado"
    ws.ChartObjects(2).Select
    Set coll = Core_SelectedCharts()
    AssertEqual 1, coll.Count, "o grafico ativo"
    AssertEqual ws.ChartObjects(2).Name, coll(1).Parent.Name
    ws.Shapes.Range(Array(ws.ChartObjects(1).Name, ws.ChartObjects(3).Name)).Select
    Set coll = Core_SelectedCharts()
    AssertEqual 2, coll.Count, "dois selecionados (Ctrl+clique)"
    AssertEqual ws.ChartObjects(1).Name, coll(1).Parent.Name
    AssertEqual ws.ChartObjects(3).Name, coll(2).Parent.Name
    Teardown wb
End Sub

' Sem selecao a "Escala automatica" desfaz a escala de todos os graficos da planilha (o par da Mesma escala sem selecao).
Public Sub Test_Core_ResetScale_NoSelection_UsesSheet_And_NoChartNotices()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Core_ClearLastNotice
    Core_SameScale                                                 ' planilha sem graficos
    AssertEqual "Mesma escala: selecione dois ou mais gráficos do ThinQQ (ou nenhum, para usar todos os gráficos da planilha).", Core_LastNotice()
    ws.Range("A1:D5").Select: Core_Insert "stacked"
    ws.Range("A1:D5").Select: Core_Insert "barclustered"
    ws.Range("A1").Select
    Core_SameScale
    AssertTrue Len(Registry_GetOption(ws.ChartObjects(1).Chart, "vmax", "")) > 0
    Core_ClearLastNotice
    Core_ResetScale
    AssertEqual "Escala automática restaurada em 2 gráfico(s).", Core_LastNotice()
    AssertEqual "", Registry_GetOption(ws.ChartObjects(1).Chart, "vmax", "")
    AssertEqual "", Registry_GetOption(ws.ChartObjects(2).Chart, "vmax", "")
    AssertTrue ws.ChartObjects(2).Chart.Axes(xlValue).MaximumScaleIsAuto
    Teardown wb
End Sub

Public Sub Test_Core_Toggles_Ind100_And_VLine_ThroughApi()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:D5").Select: Core_Insert "stacked100"
    Set cht = Core_ActiveChart()
    AssertTrue Core_GetToggle("ind100"), "ligado por padrao"
    AssertTrue Core_Supports("ind100")
    AssertFalse Core_Supports("vline")
    Core_SetToggle "ind100", False
    AssertFalse Core_GetToggle("ind100")
    AssertEqual 0, CountPrefix(cht, "tqq_ann_ind100")
    Core_SetToggle "ind100", True
    AssertTrue Core_GetToggle("ind100")
    AssertEqual 2, CountPrefix(cht, "tqq_ann_ind100")
    ws.Range("A1:D5").Select: Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    AssertFalse Core_Supports("ind100")
    Core_ClearLastNotice
    Core_SetToggle "ind100", True
    AssertEqual "O indicador 100% só existe em gráficos 100% e Mekko (%).", Core_LastNotice()
    AssertFalse Core_GetToggle("ind100")
    AssertTrue Core_Supports("vline")
    Core_SetToggle "vline", True
    AssertTrue Core_GetToggle("vline")
    AssertFalse Chart_FindSeries(cht, "tqq_vline") Is Nothing
    Core_SetToggle "vline", False
    AssertFalse Core_GetToggle("vline")
    AssertTrue Chart_FindSeries(cht, "tqq_vline") Is Nothing
    Teardown wb
End Sub

' "Alterar dados" reaplica as decoracoes SALVAS (a escolha do usuario), nunca o padrao nem o estado real: o indicador 100% desligado
' continua desligado, a linha de valor ligada volta e a escala fixa (Mesma escala) e reaplicada ao eixo do grafico novo.
Public Sub Test_Core_RebuildFromRange_ReappliesIndicatorVLineAndFixedScale()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, c As Chart, nc As Chart, c2 As Chart, coll As New Collection, skipped As Long
    Dim mx As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ' 100%: indicador desligado pelo usuario
    ws.Range("A1:D5").Select: Core_Insert "stacked100"
    Set c = Core_ActiveChart()
    Core_SetToggle "ind100", False
    AssertTrue Core_RebuildFromRange(c, ws.Range("A1:D5"))
    Set nc = Core_ActiveChart()
    AssertEqual "0", Registry_GetOption(nc, "ind100", ""), "a escolha salva (desligado) vence o padrao do Insert"
    AssertEqual 0, CountPrefix(nc, "tqq_ann_ind100")
    ' 100%: ligado (padrao) volta ligado, com as duas formas
    Core_SetToggle "ind100", True
    AssertTrue Core_RebuildFromRange(nc, ws.Range("A1:D5"))
    Set nc = Core_ActiveChart()
    AssertEqual "1", Registry_GetOption(nc, "ind100", "")
    AssertEqual 2, CountPrefix(nc, "tqq_ann_ind100")
    ' empilhado: linha de valor ligada e escala fixa persistem
    ws.Range("A1:D5").Select: Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Core_SetToggle "vline", True
    ws.Range("F1:I1").Value = Array("Ano", "A", "B", "C")
    ws.Range("F2:I2").Value = Array(2022, 100, 200, 50): ws.Range("F3:I3").Value = Array(2023, 120, 220, 60)
    Set c2 = Stacked_Insert(ws, ChartData_FromRange(ws.Range("F1:I3")), "stacked").Chart
    coll.Add c: coll.Add c2
    AssertEqual 2, Deco_SameScale(coll, skipped)
    mx = c.Axes(xlValue).MaximumScale
    AssertTrue mx >= 400, "escala do grafico maior"
    AssertTrue Core_RebuildFromRange(c, ws.Range("A1:D5"))
    Set nc = Core_ActiveChart()
    AssertEqual "1", Registry_GetOption(nc, "vline", "")
    AssertFalse Chart_FindSeries(nc, "tqq_vline") Is Nothing, "a linha de valor voltou"
    AssertFalse nc.Axes(xlValue).MaximumScaleIsAuto, "a escala fixa voltou ao eixo"
    AssertAlmost mx, nc.Axes(xlValue).MaximumScale, 0.000001
    Teardown wb
End Sub
