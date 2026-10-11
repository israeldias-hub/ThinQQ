Attribute VB_Name = "TestGantt"
Option Explicit
' Datas: 2026-03-02 e 2026-04-20 sao segundas-feiras (WEEKDAY(d,3) = 0).

Private Function Fixture(wb As Workbook, Optional withOwner As Boolean = False) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Plano"
    ws.Range("A1:C1").Value = Array("Atividade", "Início", "Fim")
    ws.Range("A2").Value = "Planejamento": ws.Range("B2").Value = DateSerial(2026, 3, 2): ws.Range("C2").Value = DateSerial(2026, 3, 13)
    ws.Range("A3").Value = "Execução": ws.Range("B3").Value = DateSerial(2026, 3, 9): ws.Range("C3").Value = DateSerial(2026, 4, 17)
    ws.Range("A4").Value = "Entrega": ws.Range("B4").Value = DateSerial(2026, 4, 20)
    ws.Range("B2:C4").NumberFormat = "dd/mm/yyyy"
    If withOwner Then
        ws.Range("D1").Value = "Responsável": ws.Range("D2:D4").Value = Application.Transpose(Array("Ana", "Bruno", "Carla"))
        Set Fixture = ChartData_FromRange(ws.Range("A1:D4"))
    Else
        Set Fixture = ChartData_FromRange(ws.Range("A1:C4"))
    End If
End Function

Private Function ValidTicks(body As Range) As Long
    Dim r As Long
    For r = 2 To GT_SCALE_SLOTS + 1
        If Not IsError(body.Cells(r, GT_COL_TICK).Value) Then ValidTicks = ValidTicks + 1
    Next
End Function

Public Sub Test_Gantt_Block_FormulasAndWeekScale()
    Dim wb As Workbook, blk As Range, body As Range
    Set wb = NewTempWorkbook()
    Set blk = Gantt_BuildBlock(wb, "gt0001", Fixture(wb)): Set body = Block_Body(blk)
    AssertEqual 19, blk.Columns.Count: AssertEqual 41, body.Rows.Count, "corpo: cabecalho + max(n, 40) (o bloco tem ainda a linha de meta)"
    AssertEqual CDbl(DateSerial(2026, 3, 2)), body.Cells(2, GT_COL_DMIN).Value: AssertEqual CDbl(DateSerial(2026, 4, 20)), body.Cells(2, GT_COL_DMAX).Value
    AssertEqual False, body.Cells(2, GT_COL_ISMS).Value: AssertEqual True, body.Cells(4, GT_COL_ISMS).Value
    AssertEqual CDbl(DateSerial(2026, 3, 2)), body.Cells(2, GT_COL_BARSTART).Value
    AssertEqual 12, body.Cells(2, GT_COL_DUR).Value, "fim inclusivo: 2 a 13 de marco"
    AssertEqual 40, body.Cells(3, GT_COL_DUR).Value
    AssertTrue IsError(body.Cells(4, GT_COL_DUR).Value), "marco nao tem barra"
    AssertTrue IsError(body.Cells(2, GT_COL_MSX).Value): AssertEqual CDbl(DateSerial(2026, 4, 20)), body.Cells(4, GT_COL_MSX).Value
    AssertEqual 2.5, body.Cells(2, GT_COL_POS).Value: AssertEqual 0.5, body.Cells(4, GT_COL_POS).Value
    AssertEqual CDbl(DateSerial(2026, 3, 14)), body.Cells(2, GT_COL_LBLX).Value, "rotulo depois do fim da barra (Fim + 1)"
    AssertEqual CDbl(DateSerial(2026, 4, 20)), body.Cells(4, GT_COL_LBLX).Value, "marco: no proprio dia"
    AssertEqual "w", body.Cells(2, GT_COL_UNIT).Value, "49 dias: semanas"
    AssertEqual CDbl(DateSerial(2026, 3, 2)), body.Cells(2, GT_COL_AXMIN).Value
    AssertEqual CDbl(DateSerial(2026, 4, 27)), body.Cells(2, GT_COL_AXMAX).Value, "segunda seguinte ao fim"
    AssertEqual 8, ValidTicks(body)
    AssertEqual CDbl(DateSerial(2026, 3, 9)), body.Cells(3, GT_COL_TICK).Value
    AssertEqual 3, body.Cells(2, GT_COL_TICKY).Value: AssertTrue IsError(body.Cells(10, GT_COL_TICKY).Value)
    CloseTemp wb
End Sub

Public Sub Test_Gantt_Insert_BarsMilestonesScale()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ms As Series, sc As Series, lbl As Series, v As Variant, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Gantt_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "gantt", Registry_ChartType(cht)
    AssertEqual xlBarStacked, Chart_FindSeries(cht, "tqq_start").ChartType
    AssertEqual msoFalse, Chart_FindSeries(cht, "tqq_start").Format.Fill.Visible
    AssertTrue cht.Axes(xlCategory).ReversePlotOrder
    AssertEqual CDbl(DateSerial(2026, 3, 2)), cht.Axes(xlValue, xlPrimary).MinimumScale
    AssertEqual CDbl(DateSerial(2026, 4, 27)), cht.Axes(xlValue, xlPrimary).MaximumScale
    Set ms = Chart_FindSeries(cht, "tqq_ms")
    AssertEqual xlSecondary, ms.AxisGroup: AssertEqual xlMarkerStyleDiamond, ms.MarkerStyle
    v = ms.XValues
    AssertEqual CDbl(DateSerial(2026, 4, 20)), v(3)
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    AssertEqual Format$(DateSerial(2026, 3, 13), "dd/mm"), lbl.Points(1).DataLabel.Text
    AssertEqual Format$(DateSerial(2026, 4, 20), "dd/mm"), lbl.Points(3).DataLabel.Text, "marco rotulado com a propria data"
    AssertEqual xlLabelPositionRight, lbl.Points(1).DataLabel.Position
    Set sc = Chart_FindSeries(cht, "tqq_scale")
    AssertEqual 40, sc.Points.Count
    AssertEqual Format$(DateSerial(2026, 3, 2), "dd/mm"), sc.Points(1).DataLabel.Text
    AssertEqual xlLabelPositionAbove, sc.Points(1).DataLabel.Position
    AssertTrue sc.HasErrorBars, "separadores verticais por barras de erro"
    AssertFalse cht.HasLegend
    AssertEqual "50", Registry_GetOption(cht, "gap"): AssertEqual "3", Registry_GetOption(cht, "n"): AssertEqual "1", Registry_GetOption(cht, "labels")
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual "dd/mm", body.Cells(2, GT_COL_TICKLBL).NumberFormat
    AssertEqual "dd/mm", body.Cells(2, GT_COL_LBLTXT).NumberFormat, "ligar os rotulos nao troca o formato de data das celulas"
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    AssertEqual 1, CalcSheet_Cleanup(wb), "limpeza com o grafico novo (auxiliares com pontos NA) remove so o orfao"
    AssertFalse CalcSheet_CleanupReadFailed(), "pontos NA dos auxiliares nao sao falha de leitura"
    Gantt_SetLabels cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl") Is Nothing: AssertEqual "0", Registry_GetOption(cht, "labels")
    CloseTemp wb
End Sub

Public Sub Test_Gantt_Geometry_Independent()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lbl As Series, ax As Axis, il As Double, it As Double, iw As Double, ih As Double, x As Double, dl As DataLabel
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Gantt_Insert(d.Source.Worksheet, d).Chart
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    il = cht.PlotArea.InsideLeft: it = cht.PlotArea.InsideTop: iw = cht.PlotArea.InsideWidth: ih = cht.PlotArea.InsideHeight
    Set ax = cht.Axes(xlValue, xlPrimary)
    x = il + iw * (CDbl(DateSerial(2026, 4, 18)) - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)      ' fim visual da barra = Fim + 1
    Set dl = lbl.Points(2).DataLabel
    AssertTrue dl.Left >= x - 2 And dl.Left <= x + 14, "rotulo da Execucao no fim da barra: " & dl.Left & " vs " & x
    AssertTrue Abs((dl.Top + dl.Height / 2) - (it + 1.5 * ih / 3)) <= ih / 6, "na faixa 2"
    x = il + iw * (CDbl(DateSerial(2026, 4, 20)) - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
    Set dl = lbl.Points(3).DataLabel
    AssertTrue dl.Left >= x - 2 And dl.Left <= x + 16, "rotulo do marco a direita do losango"
    AssertTrue Abs((dl.Top + dl.Height / 2) - (it + 2.5 * ih / 3)) <= ih / 6, "na faixa 3"
    Set dl = Chart_FindSeries(cht, "tqq_scale").Points(1).DataLabel
    AssertTrue dl.Top + dl.Height <= it + 2, "rotulo da escala acima da area de plotagem"
    AssertTrue Abs((dl.Left + dl.Width / 2) - il) <= 12, "primeira marca na borda esquerda (AxisMin)"
    AssertTrue dl.Top >= 0, "rotulo da escala dentro do grafico"
    AssertTrue cht.PlotArea.InsideTop >= 23.5, "folga acima da area de plotagem para a escala"
    CloseTemp wb
End Sub

' Review Focus 4: a unidade da escala muda com os dados e o refresh reaplica eixo e formatos.
Public Sub Test_Gantt_ScaleUnitSwitches_OnRefresh()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Gantt_Insert(d.Source.Worksheet, d).Chart
    Set body = Block_Body(Registry_Block(cht))
    d.Source.Worksheet.Range("C3").Value = DateSerial(2027, 1, 15)       ' 319 dias: meses
    AssertEqual "m", body.Cells(2, GT_COL_UNIT).Value
    Gantt_RefreshScale cht
    AssertEqual CDbl(DateSerial(2026, 3, 1)), cht.Axes(xlValue, xlPrimary).MinimumScale
    AssertEqual CDbl(DateSerial(2027, 2, 1)), cht.Axes(xlValue, xlPrimary).MaximumScale
    AssertEqual 11, ValidTicks(body), "mar/26 a jan/27"
    AssertEqual "mmm/yy", body.Cells(2, GT_COL_TICKLBL).NumberFormat
    d.Source.Worksheet.Range("C3").Value = DateSerial(2029, 6, 30)       ' > 1100 dias: trimestres
    Core_SafeRefresh cht
    AssertEqual "q", body.Cells(2, GT_COL_UNIT).Value
    AssertEqual CDbl(DateSerial(2026, 1, 1)), cht.Axes(xlValue, xlPrimary).MinimumScale
    AssertEqual CDbl(DateSerial(2029, 7, 1)), cht.Axes(xlValue, xlPrimary).MaximumScale
    AssertEqual 14, ValidTicks(body)
    AssertEqual "T1/2026", body.Cells(2, GT_COL_TICKLBL).Value
    AssertEqual "T1/2026", Chart_FindSeries(cht, "tqq_scale").Points(1).DataLabel.Text
    CloseTemp wb
End Sub

Public Sub Test_Gantt_InvalidDates_RefusedWithMessage()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, en As Long, ed As String, nm As Name
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("Atividade", "Início", "Fim")
    ws.Range("A2:C2").Value = Array("Planejamento", "2 de março", "13 de março")
    ws.Range("A3:C3").Value = Array("Execução", "depois", "")
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    On Error Resume Next
    Gantt_Insert ws, d
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 6010, en
    AssertEqual "Gantt: a coluna de início não tem datas válidas (use células de data, não texto).", ed
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nada alocado"
    Next
    ws.Range("A1:B3").Select
    Core_ClearLastNotice
    Core_Insert "gantt"
    AssertTrue InStr(Core_LastNotice(), "3 colunas") > 0, Core_LastNotice()
    Watchers_Stop
    CloseTemp wb
End Sub

Public Sub Test_Gantt_OnlyMilestones_And_SingleActivity()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("Atividade", "Início", "Fim")
    ws.Range("A2").Value = "Kick-off": ws.Range("B2").Value = DateSerial(2026, 5, 6)
    Set d = ChartData_FromRange(ws.Range("A1:C2"))
    Set cht = Gantt_Insert(ws, d).Chart
    AssertEqual "1", Registry_GetOption(cht, "n")
    AssertEqual 1, Chart_FindSeries(cht, "tqq_ms").Points.Count
    AssertEqual 1, cht.Axes(xlValue, xlSecondary).MaximumScale
    AssertEqual CDbl(DateSerial(2026, 5, 4)), cht.Axes(xlValue, xlPrimary).MinimumScale, "segunda anterior"
    AssertEqual CDbl(DateSerial(2026, 5, 11)), cht.Axes(xlValue, xlPrimary).MaximumScale
    AssertEqual Format$(DateSerial(2026, 5, 6), "dd/mm"), Chart_FindSeries(cht, "tqq_lbl").Points(1).DataLabel.Text
    CloseTemp wb
End Sub

Public Sub Test_Gantt_FourthColumn_AsLabel()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb, True)
    Set cht = Gantt_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "Ana", Chart_FindSeries(cht, "tqq_lbl").Points(1).DataLabel.Text
    AssertEqual "Carla", Chart_FindSeries(cht, "tqq_lbl").Points(3).DataLabel.Text
    CloseTemp wb
End Sub

Public Sub Test_Gantt_Bracket_Annotation()
    Dim wb As Workbook, d As clsChartData, cht As Chart, br As Shape, lbl As Shape, ax As Axis, il As Double, iw As Double, en As Long, ed As String, other As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Gantt_Insert(d.Source.Worksheet, d).Chart
    Set br = Ann_AddBracket(cht, 1, 2)
    AssertEqual "tqq_ann_bracket_1", br.Name
    AssertEqual "kind=bracket;from=1;to=2", br.AlternativeText
    AssertTrue Not cht.Shapes("tqq_ann_bracket_1_t1") Is Nothing And Not cht.Shapes("tqq_ann_bracket_1_t2") Is Nothing
    Set lbl = cht.Shapes("tqq_ann_bracket_1_lbl")
    AssertEqual Format$(DateSerial(2026, 3, 2), "dd/mm") & " - " & Format$(DateSerial(2026, 4, 17), "dd/mm"), lbl.TextFrame2.TextRange.Text
    il = cht.PlotArea.InsideLeft: iw = cht.PlotArea.InsideWidth: Set ax = cht.Axes(xlValue, xlPrimary)
    AssertAlmost il + iw * (CDbl(DateSerial(2026, 3, 2)) - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale), br.Left, 1
    AssertAlmost iw * (CDbl(DateSerial(2026, 4, 18)) - CDbl(DateSerial(2026, 3, 2))) / (ax.MaximumScale - ax.MinimumScale), br.Width, 1, "ate o fim visual da barra (Fim + 1)"
    AssertTrue br.Top < cht.PlotArea.InsideTop - 10, "colchete acima da area de plotagem"
    AssertTrue lbl.Top + lbl.Height <= br.Top + 0.5, "rotulo acima do colchete"
    AssertTrue lbl.Top >= 0
    Ann_AddBracket cht, 3, 3
    AssertTrue cht.Shapes("tqq_ann_bracket_2").Top < br.Top, "segundo colchete empilha acima"
    Ann_RemoveAll cht
    AssertEqual 0, cht.Shapes.Count
    d.Source.Worksheet.Range("F1:G1").Value = Array("Ano", "V"): d.Source.Worksheet.Range("F2:G2").Value = Array(2022, 1): d.Source.Worksheet.Range("F3:G3").Value = Array(2023, 2)
    Set other = Stacked_Insert(d.Source.Worksheet, ChartData_FromRange(d.Source.Worksheet.Range("F1:G3")), "stacked").Chart
    On Error Resume Next
    Ann_AddBracket other, 1, 2
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 7001, en: AssertEqual "O colchete só está disponível em gráficos Gantt.", ed
    CloseTemp wb
End Sub

Public Sub Test_Core_Insert_Gantt_And_Rollback()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, n As Long, nm As Name
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:C4").Select
    Core_Insert "gantt"
    Set cht = Core_ActiveChart()
    AssertEqual "gantt", Registry_ChartType(cht)
    AssertTrue Core_Supports("bracket"): AssertFalse Core_Supports("numfmt"): AssertTrue Core_Supports("gap")
    Core_SetOption "gap", "80"
    AssertEqual 80, cht.ChartGroups(1).GapWidth
    Core_AddAnnotation "bracket", 1, 3
    AssertTrue Not cht.Shapes("tqq_ann_bracket_1") Is Nothing
    Core_AddAnnotation "bracket", 2, 2
    AssertTrue Not cht.Shapes("tqq_ann_bracket_2") Is Nothing, "colchete sobre uma unica atividade (de = ate) sem reabrir o formulario"
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(127, 127, 127), Chart_FindSeries(cht, "tqq_dur").Format.Fill.ForeColor.RGB, "Gantt_ApplyColors pela API: tqq_dur = pal(1)"
    Core_SetToggle "labels", False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl") Is Nothing
    Core_ClearLastNotice
    Core_Refresh
    AssertEqual 0, Core_NoticeCount(), "sem aviso de tabela mudada"
    ws.ChartObjects(1).Delete: CalcSheet_Cleanup wb
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Gantt_Insert ws, d
    n = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, n
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    Watchers_Stop
    CloseTemp wb
End Sub

' --- Alem do brief: sondas da Tarefa 10 ---

' Le a geometria real dos rotulos depois de um desenho (o Excel invisivel nao redesenha sozinho; ver TestMekko).
Private Sub ForceLayout(cht As Chart)
    cht.Export TempDir() & "\layout.png"
End Sub

' Regra de tres sobre a area interna e a escala REAIS do grafico (independe das formulas do bloco): x de um numero de serie de data.
Private Function XOf(cht As Chart, d As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue, xlPrimary)
    XOf = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (d - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
End Function

' Escala de trimestres na insercao: com mais marcas que espaco os rotulos que aparecem nao se sobrepoem (geometria real), so marcas com data tem
' rotulo e atividade sem inicio fica sem rotulo; marcador na cor da paleta. Trimestres (e anos) -> meses: o rotulo mostra o mes da localidade
' como a celula (medido: sem passar por um formato de data, de General para mmm/yy o rotulo ja ligado mostrava "Jan/26"/"Feb/26" com a celula
' em "jan/26"); as marcas que deixam de existir perdem o rotulo e as que passam a existir ganham.
Public Sub Test_Gantt_MonthScale_LocaleNames_Thinning_NoStrayLabels()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, sc As Series, body As Range, r As Long, labeled As Long, prevRight As Double, dl As DataLabel
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("Atividade", "Início", "Fim")
    ws.Range("A2:A4").Value = Application.Transpose(Array("Fase 1", "Fase 2", "Fase 3"))
    ws.Range("B2").Value = DateSerial(2026, 1, 12): ws.Range("C2").Value = DateSerial(2026, 3, 20)
    ws.Range("B3").Value = DateSerial(2026, 3, 2): ws.Range("C3").Value = DateSerial(2029, 6, 30)
    ws.Range("C4").Value = DateSerial(2026, 5, 15)
    Set cht = Gantt_Insert(ws, ChartData_FromRange(ws.Range("A1:C4"))).Chart
    AssertTrue Geom_ChartIsHorizontal(cht), "Gantt e horizontal (categorias na vertical)"
    Set body = Block_Body(Registry_Block(cht))
    Set sc = Chart_FindSeries(cht, "tqq_scale")
    AssertEqual "q", body.Cells(2, GT_COL_UNIT).Value, "14 trimestres"
    AssertEqual 14, ValidTicks(body)
    AssertFalse Chart_FindSeries(cht, "tqq_lbl").Points(3).HasDataLabel, "fim sem inicio: sem barra, sem marco e sem rotulo solto"
    AssertEqual Style_Palette("padrao")(0), Chart_FindSeries(cht, "tqq_ms").MarkerBackgroundColor, "losango na cor pal(0)"
    ws.Range("C3").Value = DateSerial(2026, 6, 30)
    Core_SafeRefresh cht
    AssertEqual "m", body.Cells(2, GT_COL_UNIT).Value, "169 dias: meses"
    AssertEqual Format$(DateSerial(2026, 1, 1), "mmm/yy"), sc.Points(1).DataLabel.Text, "trimestres -> meses: mes da localidade, como a celula"
    AssertEqual Format$(DateSerial(2026, 2, 1), "mmm/yy"), sc.Points(2).DataLabel.Text
    AssertEqual Format$(DateSerial(2026, 6, 1), "mmm/yy"), sc.Points(6).DataLabel.Text
    For r = 7 To GT_SCALE_SLOTS
        AssertFalse sc.Points(r).HasDataLabel, "marca que deixou de existir perde o rotulo: ponto " & r
    Next
    ws.Range("C3").Value = DateSerial(2029, 6, 30)
    Core_SafeRefresh cht
    AssertEqual "q", body.Cells(2, GT_COL_UNIT).Value, "de volta aos trimestres"
    ForceLayout cht
    prevRight = -1000
    For r = 1 To GT_SCALE_SLOTS
        If sc.Points(r).HasDataLabel Then
            AssertFalse IsError(body.Cells(r + 1, GT_COL_TICK).Value), "rotulo so em marca com data: ponto " & r
            labeled = labeled + 1
            Set dl = sc.Points(r).DataLabel
            AssertTrue dl.Left >= prevRight - 0.5, "rotulos da escala sem sobreposicao: ponto " & r & " em " & dl.Left & ", anterior termina em " & prevRight
            AssertTrue dl.Left + dl.Width <= cht.ChartArea.Width + 0.5, "rotulo da escala dentro do grafico: ponto " & r
            prevRight = dl.Left + dl.Width
        End If
    Next
    AssertTrue sc.Points(1).HasDataLabel, "a primeira marca sempre tem rotulo"
    AssertTrue labeled >= 5 And labeled < 14, "rotulos espacados: " & labeled & " de 14"
    Gantt_ApplyColors cht, "cinza"
    AssertEqual RGB(64, 64, 64), Chart_FindSeries(cht, "tqq_ms").MarkerForegroundColor
    ws.Range("C3").Value = DateSerial(2036, 6, 30)
    Core_SafeRefresh cht
    AssertEqual "y", body.Cells(2, GT_COL_UNIT).Value
    AssertEqual "2026", sc.Points(1).DataLabel.Text
    AssertEqual "2036", sc.Points(11).DataLabel.Text, "marca nova ganha rotulo no refresh"
    ws.Range("C3").Value = DateSerial(2026, 6, 30)
    Core_SafeRefresh cht
    AssertEqual Format$(DateSerial(2026, 1, 1), "mmm/yy"), sc.Points(1).DataLabel.Text, "anos -> meses: mes da localidade"
    AssertEqual Format$(DateSerial(2026, 2, 1), "mmm/yy"), sc.Points(2).DataLabel.Text
    CloseTemp wb
End Sub

' Folgas e opcoes pelo Core: o rotulo de uma barra que termina no fim do eixo fica depois da barra e dentro do grafico (sem a folga o Excel o
' empurra para cima da barra); as folgas voltam depois do layout automatico (Chart_EnsureRoomAfterChange); legenda e linha de valor recusadas
' (7002, aviso); totais nao gravam nada; Alterar dados reaplica as opcoes salvas; falha depois do grafico desfaz tudo.
Public Sub Test_Gantt_RightRoom_LayoutReset_Refusals_RebuildAndRollback()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, dl As DataLabel, ax As Axis, x As Double, n As Long, nm As Name, charts As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("C3").Value = DateSerial(2026, 4, 26)              ' domingo: Fim + 1 = 27/04 = fim do eixo semanal
    ws.Range("A1:C4").Select
    Core_Insert "gantt"
    Set cht = Core_ActiveChart()
    Set ax = cht.Axes(xlValue, xlPrimary)
    AssertEqual CDbl(DateSerial(2026, 4, 27)), ax.MaximumScale
    ForceLayout cht
    Set dl = Chart_FindSeries(cht, "tqq_lbl").Points(2).DataLabel
    x = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (CDbl(DateSerial(2026, 4, 27)) - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
    AssertTrue dl.Left >= x - 2, "rotulo depois do fim da barra: " & dl.Left & " vs " & x
    AssertTrue dl.Left + dl.Width <= cht.ChartArea.Width + 0.5, "rotulo dentro do grafico"
    Core_AddAnnotation "bracket", 3, 3
    x = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (CDbl(DateSerial(2026, 4, 20)) - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
    AssertEqual msoTrue, cht.Shapes("tqq_ann_bracket_1").Visible, "colchete sobre um marco so (de = ate) visivel"
    AssertAlmost x, cht.Shapes("tqq_ann_bracket_1").Left, 1, "no dia do marco"
    AssertEqual Format$(DateSerial(2026, 4, 20), "dd/mm"), cht.Shapes("tqq_ann_bracket_1_lbl").TextFrame2.TextRange.Text, "um dia: a data uma vez"
    AssertTrue cht.PlotArea.InsideTop >= 41.5, "folga do colchete"
    cht.PlotArea.Position = xlChartElementPositionAutomatic
    Chart_EnsureRoomAfterChange cht
    AssertTrue cht.PlotArea.InsideTop >= 41.5, "folga do colchete e da escala refeita depois do layout automatico"
    ForceLayout cht
    Set dl = Chart_FindSeries(cht, "tqq_lbl").Points(2).DataLabel
    x = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth
    AssertTrue dl.Left >= x - 2 And dl.Left + dl.Width <= cht.ChartArea.Width + 0.5, "folga a direita refeita: " & dl.Left & " vs " & x
    Core_SetOption "numfmt", "dec1"
    AssertEqual "dd/mm", Block_Body(Registry_Block(cht)).Cells(2, GT_COL_LBLTXT).NumberFormat, "numfmt nao se aplica ao Gantt"
    AssertEqual Format$(DateSerial(2026, 3, 13), "dd/mm"), Chart_FindSeries(cht, "tqq_lbl").Points(1).DataLabel.Text, "rotulo de data intacto"
    AssertEqual Format$(DateSerial(2026, 3, 2), "dd/mm"), Chart_FindSeries(cht, "tqq_scale").Points(1).DataLabel.Text
    Core_ClearLastNotice
    Core_SetOption "legend", "none"
    AssertEqual "Legenda não se aplica a este tipo de gráfico.", Core_LastNotice()
    Core_ClearLastNotice
    Core_SetToggle "vline", True
    AssertEqual "A linha de valor não está disponível para este tipo de gráfico.", Core_LastNotice()
    Core_SetToggle "totals", True
    Core_SetToggle "connectors", True
    AssertEqual "", Registry_GetOption(cht, "totals") & Registry_GetOption(cht, "connectors"), "totais e conectores nao existem no Gantt"
    Core_SetOption "palette", "cinza"
    Core_SetOption "gap", "120"
    Core_SetToggle "labels", False
    Core_RebuildFromRange cht, ws.Range("A1:C3")
    Set cht = Core_ActiveChart()
    AssertEqual "gantt", Registry_ChartType(cht)
    AssertEqual "2", Registry_GetOption(cht, "n"): AssertEqual "0", Registry_GetOption(cht, "labels")
    AssertTrue Chart_FindSeries(cht, "tqq_lbl") Is Nothing, "rotulos desligados reaplicados"
    AssertEqual RGB(127, 127, 127), Chart_FindSeries(cht, "tqq_dur").Format.Fill.ForeColor.RGB, "paleta salva reaplicada"
    AssertEqual RGB(64, 64, 64), Chart_FindSeries(cht, "tqq_ms").MarkerBackgroundColor
    AssertEqual 120, cht.ChartGroups(1).GapWidth
    AssertFalse cht.HasLegend
    charts = ws.ChartObjects.Count
    Util_FailAfterChartForTest = True
    On Error Resume Next
    Gantt_Insert ws, d
    n = Err.Number
    On Error GoTo 0
    Util_FailAfterChartForTest = False
    AssertEqual vbObjectError + 9998, n
    AssertEqual charts, ws.ChartObjects.Count, "grafico da falha apagado"
    CalcSheet_Cleanup wb
    AssertFalse CalcSheet_CleanupReadFailed()
    n = 0
    For Each nm In wb.Names
        If Left$(nm.Name, 4) = "tqq_" Then n = n + 1
    Next
    AssertEqual 1, n, "so o bloco do grafico reconstruido"
    Watchers_Stop
    CloseTemp wb
End Sub

' Barras e marco no lugar das datas (regra de tres sobre o grafico desenhado, sem as formulas do bloco): cada barra comeca no inicio e termina em
' Fim + 1 (cobre o dia do fim), na faixa da sua atividade; a parte invisivel termina no inicio; o losango fica no dia do marco, no meio da faixa 3.
' Dois colchetes empilhados: a linha do 1o 25 pt acima da area interna e a do 2o 20 pt acima dela; os tiques do 1o acabam acima dos rotulos da
' escala e os do 2o acima do rotulo do 1o; a folga dos dois volta no refresh depois do layout automatico.
Public Sub Test_Gantt_Bars_OnDates_RuleOfThree_StackedBrackets()
    Dim wb As Workbook, d As clsChartData, cht As Chart, dur As Series, st As Series, ms As Series, i As Long, it As Double, ih As Double
    Dim starts As Variant, ends As Variant, b1 As Shape, b2 As Shape, l1 As Shape, l2 As Shape, tk As Shape, sl As DataLabel, nmv As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Gantt_Insert(d.Source.Worksheet, d).Chart
    ForceLayout cht
    it = cht.PlotArea.InsideTop: ih = cht.PlotArea.InsideHeight
    starts = Array(DateSerial(2026, 3, 2), DateSerial(2026, 3, 9))
    ends = Array(DateSerial(2026, 3, 13), DateSerial(2026, 4, 17))
    Set dur = Chart_FindSeries(cht, "tqq_dur"): Set st = Chart_FindSeries(cht, "tqq_start"): Set ms = Chart_FindSeries(cht, "tqq_ms")
    For i = 1 To 2
        AssertAlmost XOf(cht, CDbl(starts(i - 1))), dur.Points(i).Left, 1, "barra " & i & " comeca no inicio"
        AssertAlmost XOf(cht, CDbl(ends(i - 1)) + 1), dur.Points(i).Left + dur.Points(i).Width, 1, "barra " & i & " termina em Fim + 1"
        AssertAlmost it + (i - 0.5) * ih / 3, dur.Points(i).Top + dur.Points(i).Height / 2, 1, "barra " & i & " na faixa " & i
    Next
    AssertAlmost XOf(cht, CDbl(DateSerial(2026, 3, 9))), st.Points(2).Left + st.Points(2).Width, 1, "a parte invisivel termina no inicio"
    AssertAlmost XOf(cht, CDbl(DateSerial(2026, 4, 20))), ms.Points(3).Left + ms.Points(3).Width / 2, 1.5, "losango no dia do marco"
    AssertAlmost it + 2.5 * ih / 3, ms.Points(3).Top + ms.Points(3).Height / 2, 1.5, "losango na faixa 3"
    Ann_AddBracket cht, 1, 2
    Ann_AddBracket cht, 2, 3
    ForceLayout cht
    it = cht.PlotArea.InsideTop
    AssertTrue it >= 61.5, "folga de dois colchetes: " & it
    Set b1 = cht.Shapes("tqq_ann_bracket_1"): Set b2 = cht.Shapes("tqq_ann_bracket_2")
    Set l1 = cht.Shapes("tqq_ann_bracket_1_lbl"): Set l2 = cht.Shapes("tqq_ann_bracket_2_lbl")
    AssertAlmost it - 25, b1.Top, 0.5, "colchete 1: 25 pt acima da area interna"
    AssertAlmost it - 45, b2.Top, 0.5, "colchete 2: 20 pt acima do 1"
    Set sl = Chart_FindSeries(cht, "tqq_scale").Points(1).DataLabel
    For Each nmv In Array("tqq_ann_bracket_1_t1", "tqq_ann_bracket_1_t2")
        Set tk = cht.Shapes(CStr(nmv))
        AssertTrue tk.Top + tk.Height <= sl.Top + 0.5, CStr(nmv) & " acima dos rotulos da escala: " & (tk.Top + tk.Height) & " vs " & sl.Top
    Next
    For Each nmv In Array("tqq_ann_bracket_2_t1", "tqq_ann_bracket_2_t2")
        Set tk = cht.Shapes(CStr(nmv))
        AssertTrue tk.Top + tk.Height <= l1.Top + 0.5, CStr(nmv) & " acima do rotulo do colchete 1: " & (tk.Top + tk.Height) & " vs " & l1.Top
    Next
    AssertTrue l1.Top + l1.Height <= b1.Top + 0.5 And l2.Top + l2.Height <= b2.Top + 0.5, "rotulos acima das linhas"
    AssertTrue l2.Top >= 0, "rotulo de cima dentro do grafico"
    cht.PlotArea.Position = xlChartElementPositionAutomatic
    Core_SafeRefresh cht
    AssertTrue cht.PlotArea.InsideTop >= 61.5, "folga dos colchetes refeita pelo refresh: " & cht.PlotArea.InsideTop
    AssertAlmost cht.PlotArea.InsideTop - 45, b2.Top, 0.5, "colchete reposicionado"
    Watchers_Stop
    CloseTemp wb
End Sub

' Pasta no sistema de datas 1904 (Workbook.Date1904, ligado ANTES de escrever as datas): celulas e eixo usam numeros de serie 1462 dias menores; os
' rotulos (texto das celulas) e o rotulo do colchete (escrito pelo VBA a partir das celulas) mostram as datas digitadas.
Public Sub Test_Gantt_Date1904_BracketAndLabels()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook()
    wb.Date1904 = True
    Set d = Fixture(wb)
    Set cht = Gantt_Insert(d.Source.Worksheet, d).Chart
    AssertEqual CDbl(DateSerial(2026, 3, 2)) - 1462, cht.Axes(xlValue, xlPrimary).MinimumScale, "eixo no sistema 1904 (segunda-feira 02/03)"
    AssertEqual Format$(DateSerial(2026, 3, 13), "dd/mm"), Chart_FindSeries(cht, "tqq_lbl").Points(1).DataLabel.Text
    AssertEqual Format$(DateSerial(2026, 3, 2), "dd/mm"), Chart_FindSeries(cht, "tqq_scale").Points(1).DataLabel.Text
    Ann_AddBracket cht, 1, 2
    AssertEqual Format$(DateSerial(2026, 3, 2), "dd/mm") & " - " & Format$(DateSerial(2026, 4, 17), "dd/mm"), cht.Shapes("tqq_ann_bracket_1_lbl").TextFrame2.TextRange.Text
    Ann_AddBracket cht, 3, 3
    AssertEqual Format$(DateSerial(2026, 4, 20), "dd/mm"), cht.Shapes("tqq_ann_bracket_2_lbl").TextFrame2.TextRange.Text
    AssertAlmost XOf(cht, CDbl(DateSerial(2026, 3, 2)) - 1462), cht.Shapes("tqq_ann_bracket_1").Left, 1, "colchete no inicio"
    CloseTemp wb
End Sub

' Limite trimestre/ano: ate 3500 dias a escala de trimestres cabe nas 40 marcas (pior caso, comecando no ultimo dia de um trimestre: 40 trimestres);
' 3501 dias ja e anos; 3690 dias (trimestres pelo limite antigo de 3700: 42 marcas) sao anos. Semana com hora no inicio: marcas no dia inteiro.
Public Sub Test_Gantt_Block_QuarterScaleFitsSlots_WeekIgnoresTime()
    Dim wb As Workbook, ws As Worksheet, body As Range, d0 As Date
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("Atividade", "Início", "Fim")
    d0 = DateSerial(2016, 3, 31)
    ws.Range("A2").Value = "X": ws.Range("B2").Value = d0: ws.Range("C2").Value = d0 + 3500
    Set body = Block_Body(Gantt_BuildBlock(wb, "gt0002", ChartData_FromRange(ws.Range("A1:C2"))))
    AssertEqual "q", body.Cells(2, GT_COL_UNIT).Value, "3500 dias: trimestres"
    AssertEqual 40, DateDiff("q", CDate(body.Cells(2, GT_COL_AXMIN).Value), CDate(body.Cells(2, GT_COL_AXMAX).Value)), "o eixo cobre 40 trimestres"
    AssertEqual 40, ValidTicks(body), "uma marca por trimestre, nas 40 posicoes"
    AssertEqual CDbl(DateSerial(2025, 10, 1)), body.Cells(GT_SCALE_SLOTS + 1, GT_COL_TICK).Value, "a ultima posicao e o ultimo trimestre"
    ws.Range("C2").Value = d0 + 3501
    AssertEqual "y", body.Cells(2, GT_COL_UNIT).Value, "3501 dias: anos"
    ws.Range("C2").Value = d0 + 3690
    AssertEqual "y", body.Cells(2, GT_COL_UNIT).Value, "3690 dias: 42 trimestres nao caberiam"
    AssertEqual 11, ValidTicks(body), "2016 a 2026"
    ws.Range("B2").Value = DateSerial(2026, 3, 4) + 0.75: ws.Range("C2").Value = DateSerial(2026, 3, 20)
    AssertEqual "w", body.Cells(2, GT_COL_UNIT).Value
    AssertEqual CDbl(DateSerial(2026, 3, 2)), body.Cells(2, GT_COL_AXMIN).Value, "segunda-feira a 0h, mesmo com hora no inicio"
    AssertEqual CDbl(DateSerial(2026, 3, 9)), body.Cells(3, GT_COL_TICK).Value
    CloseTemp wb
End Sub
