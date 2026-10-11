Attribute VB_Name = "TestAnnotations"
Option Explicit
' Nota: os textos com percentual esperam virgula decimal (+17,0%) porque Format$ segue a localidade do
' sistema e o Excel desta maquina e pt-BR; em maquina en-US ajuste os testes para ponto.

Private Function StackedFixture(wb As Workbook) As Chart
    Dim ws As Worksheet, d As clsChartData
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 100, 50)
    ws.Range("A3:C3").Value = Array(2023, 110, 60)
    ws.Range("A4:C4").Value = Array(2024, 120, 70)
    ws.Range("A5:C5").Value = Array(2025, 150, 90)
    Set d = ChartData_FromRange(ws.Range("A1:C5"))
    Set StackedFixture = Stacked_Insert(ws, d, "stacked").Chart
End Function

Private Function WaterfallFixture(wb As Workbook) As Chart
    Dim ws As Worksheet, d As clsChartData
    Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Início", 100)
    ws.Range("A3:B3").Value = Array("A", -30)
    ws.Range("A4:B4").Value = Array("B", 50)
    ws.Range("A5:B5").Value = Array("Fim", "e")
    Set d = ChartData_FromRange(ws.Range("A1:B5"))
    Set WaterfallFixture = Waterfall_Insert(ws, d).Chart
End Function

Public Sub Test_Ann_TryCagr()
    Dim r As Double
    AssertTrue Ann_TryCagr(100, 200, 2, r)
    AssertAlmost 0.414213562, r
    AssertFalse Ann_TryCagr(0, 200, 2, r)
    AssertFalse Ann_TryCagr(100, -50, 2, r)
    AssertFalse Ann_TryCagr(100, 200, 0, r)
End Sub

Public Sub Test_Ann_ParseParam()
    AssertEqual "cagr", Ann_ParseParam("kind=cagr;from=1;to=4", "kind")
    AssertEqual "4", Ann_ParseParam("kind=cagr;from=1;to=4", "to")
    AssertEqual "", Ann_ParseParam("kind=cagr;from=1;to=4", "x")
End Sub

Public Sub Test_Ann_CategoryData_Stacked()
    Dim wb As Workbook, cht As Chart, v() As Double, t() As Double
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    Ann_GetCategoryData cht, v, t
    AssertEqual 4, UBound(v)
    AssertAlmost 150, v(1): AssertAlmost 240, v(4)
    AssertAlmost 150, t(1)
    CloseTemp wb
End Sub

Public Sub Test_Ann_CategoryData_Waterfall()
    Dim wb As Workbook, cht As Chart, v() As Double, t() As Double
    Set wb = NewTempWorkbook(): Set cht = WaterfallFixture(wb)
    Ann_GetCategoryData cht, v, t
    AssertAlmost 100, v(1): AssertAlmost 70, v(2): AssertAlmost 120, v(3): AssertAlmost 120, v(4)
    CloseTemp wb
End Sub

Public Sub Test_Ann_AddCagr_CreatesShapesAndLabel()
    Dim wb As Workbook, cht As Chart, ln As Shape, lbl As Shape
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    Set ln = Ann_AddCagr(cht, 1, 4)
    AssertEqual "tqq_ann_cagr_1", ln.Name
    AssertEqual "kind=cagr;from=1;to=4", ln.AlternativeText
    Set lbl = cht.Shapes("tqq_ann_cagr_1_lbl")
    AssertEqual "CAGR +17,0%", lbl.TextFrame2.TextRange.Text      ' (240/150)^(1/3)-1 = 16,96%
    AssertAlmost Geom_ChartCategoryCenterX(cht, 1), ln.Left, 0.5
    AssertAlmost Geom_ChartCategoryCenterX(cht, 4) - Geom_ChartCategoryCenterX(cht, 1), ln.Width, 0.5
    Set ln = Ann_AddCagr(cht, 2, 3)
    AssertEqual "tqq_ann_cagr_2", ln.Name
    CloseTemp wb
End Sub

Public Sub Test_Ann_AddDifference_LabelText()
    Dim wb As Workbook, cht As Chart, ln As Shape
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    Set ln = Ann_AddDifference(cht, 1, 4)
    AssertEqual "+90 (+60%)", cht.Shapes("tqq_ann_diff_1_lbl").TextFrame2.TextRange.Text
    AssertTrue Not cht.Shapes("tqq_ann_diff_1_lvl") Is Nothing
    AssertAlmost 0, ln.Width, 0.01, "seta de diferenca e vertical"
    CloseTemp wb
End Sub

Public Sub Test_Ann_Refresh_FollowsData()
    Dim wb As Workbook, cht As Chart, lbl As Shape
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    Ann_AddCagr cht, 1, 4
    wb.Worksheets("Dados").Range("B5").Value = 250          ' total 2025 = 340
    Ann_Refresh cht
    Set lbl = cht.Shapes("tqq_ann_cagr_1_lbl")
    AssertEqual "CAGR +31,4%", lbl.TextFrame2.TextRange.Text   ' (340/150)^(1/3)-1
    CloseTemp wb
End Sub

Public Sub Test_Ann_WaterfallConnectors()
    Dim wb As Workbook, cht As Chart, i As Long, n As Long
    Set wb = NewTempWorkbook(): Set cht = WaterfallFixture(wb)
    AssertEqual "1", Registry_GetOption(cht, "connectors")
    For i = 1 To cht.Shapes.Count
        If Left$(cht.Shapes(i).Name, 13) = "tqq_ann_conn_" Then n = n + 1
    Next
    AssertEqual 3, n
    AssertAlmost Geom_ChartValueToY(cht, 100), cht.Shapes("tqq_ann_conn_1").Top, 0.5
    Ann_SetWaterfallConnectors cht, False
    n = 0
    For i = 1 To cht.Shapes.Count
        If Left$(cht.Shapes(i).Name, 13) = "tqq_ann_conn_" Then n = n + 1
    Next
    AssertEqual 0, n
    CloseTemp wb
End Sub

Public Sub Test_Ann_RemoveAll_KeepsConnectors()
    Dim wb As Workbook, cht As Chart, i As Long, conn As Long, others As Long
    Set wb = NewTempWorkbook(): Set cht = WaterfallFixture(wb)
    Ann_AddCagr cht, 1, 4
    Ann_AddDifference cht, 1, 4
    Ann_RemoveAll cht
    For i = 1 To cht.Shapes.Count
        If Left$(cht.Shapes(i).Name, 13) = "tqq_ann_conn_" Then conn = conn + 1 Else others = others + 1
    Next
    AssertEqual 3, conn
    AssertEqual 0, others
    Ann_RemoveAll cht, True
    AssertEqual 0, cht.Shapes.Count
    CloseTemp wb
End Sub

' ---------- geometria independente (nao usa Geom_* para o valor esperado) ----------

Public Sub Test_Ann_Geometry_Independent_Stacked()
    Dim wb As Workbook, cht As Chart, ln As Shape, lbl As Shape, il As Double, it As Double, iw As Double, ih As Double
    Dim ax As Axis, yTop As Double, png As String
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    Set ln = Ann_AddCagr(cht, 1, 4)
    Set lbl = cht.Shapes("tqq_ann_cagr_1_lbl")
    il = cht.PlotArea.InsideLeft: it = cht.PlotArea.InsideTop: iw = cht.PlotArea.InsideWidth: ih = cht.PlotArea.InsideHeight
    AssertTrue iw > 50 And ih > 50, "area interna deve ter tamanho real (layout calculado): " & iw & "x" & ih
    AssertTrue ln.Left > il And ln.Left + ln.Width < il + iw, "linha do CAGR dentro da area de plotagem em x"
    ' centros de categoria: 1/8 e 7/8 da largura interna (4 categorias)
    AssertAlmost il + iw / 8, ln.Left, 1
    AssertAlmost il + iw * 7 / 8, ln.Left + ln.Width, 1
    ' eixo: valor 240 (maior total) mapeado por regra de tres independente
    Set ax = cht.Axes(xlValue)
    yTop = it + ih * (ax.MaximumScale - 240) / (ax.MaximumScale - ax.MinimumScale)
    AssertTrue ln.Top <= yTop - 5, "extremo alto da linha do CAGR acima do topo da barra mais alta (y menor)"
    AssertTrue ln.Top >= 0, "linha dentro do grafico"
    AssertTrue lbl.Top + lbl.Height <= ln.Top, "rotulo acima do extremo alto da linha"
    ' seta aponta para o fim (ultima categoria, mais alta): ponta de baixo no inicio => linha sobe para a direita
    AssertTrue ln.VerticalFlip = msoTrue, "linha vai de baixo-esquerda para cima-direita: VerticalFlip"
    AssertFalse ln.HorizontalFlip = msoTrue, "sem HorizontalFlip"
    ' renderizacao nao pode falhar
    png = TempDir() & "\ann_stacked.png"
    AssertTrue cht.Export(png), "Chart.Export deve funcionar"
    AssertTrue Len(Dir$(png)) > 0, "PNG exportado existe"
    CloseTemp wb
End Sub

Public Sub Test_Ann_Geometry_Independent_Waterfall()
    Dim wb As Workbook, cht As Chart, c As Shape, il As Double, it As Double, iw As Double, ih As Double
    Dim ax As Axis, yEnd As Double
    Set wb = NewTempWorkbook(): Set cht = WaterfallFixture(wb)
    il = cht.PlotArea.InsideLeft: it = cht.PlotArea.InsideTop: iw = cht.PlotArea.InsideWidth: ih = cht.PlotArea.InsideHeight
    Set ax = cht.Axes(xlValue)
    yEnd = it + ih * (ax.MaximumScale - 100) / (ax.MaximumScale - ax.MinimumScale)    ' End da 1a barra = 100
    Set c = cht.Shapes("tqq_ann_conn_1")
    AssertAlmost yEnd, c.Top, 1
    AssertAlmost 0, c.Height, 0.01, "conector horizontal"
    AssertTrue c.Top >= it And c.Top <= it + ih, "conector dentro da area de plotagem"
    AssertTrue c.Left > il + iw / 8 And c.Left + c.Width < il + iw * 3 / 8, "conector entre as barras 1 e 2"
    CloseTemp wb
End Sub

Public Sub Test_Ann_MoveLine_ArrowAtEnd()
    Dim wb As Workbook, cht As Chart, ln As Shape
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    ' fim a direita/acima do inicio (CAGR crescente): VerticalFlip
    Set ln = Ann_AddCagr(cht, 1, 4)
    AssertTrue ln.VerticalFlip = msoTrue
    ' dado decrescente: a seta desce para a direita, sem flips
    wb.Worksheets("Dados").Range("B5").Value = 0: wb.Worksheets("Dados").Range("C5").Value = 10
    Ann_Refresh cht
    AssertFalse ln.VerticalFlip = msoTrue, "linha descendente nao tem flip vertical"
    AssertFalse ln.HorizontalFlip = msoTrue
    CloseTemp wb
End Sub

' ---------- rodada de correcoes ----------

Private Function KindFixture(wb As Workbook, kind As String, Optional rows As Long = 4) As Chart
    Dim ws As Worksheet, d As clsChartData, i As Long
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    For i = 1 To rows
        ws.Range("A" & (i + 1) & ":C" & (i + 1)).Value = Array(2021 + i, 100 + 10 * i, 50 + 10 * i)
    Next
    Set d = ChartData_FromRange(ws.Range("A1:C" & (rows + 1)))
    Set KindFixture = Stacked_Insert(ws, d, kind).Chart
End Function

Private Function CountPrefix(cht As Chart, p As String) As Long
    Dim i As Long
    For i = 1 To cht.Shapes.Count
        If Left$(cht.Shapes(i).Name, Len(p)) = p Then CountPrefix = CountPrefix + 1
    Next
End Function

Public Sub Test_Ann_Stacked100_DifferenceRaises7001()
    Dim wb As Workbook, cht As Chart, en As Long
    Set wb = NewTempWorkbook(): Set cht = KindFixture(wb, "stacked100")
    On Error Resume Next
    Ann_AddDifference cht, 1, 4
    en = Err.Number
    On Error GoTo 0
    AssertEqual vbObjectError + 7001, en
    AssertEqual 0, CountPrefix(cht, "tqq_ann_diff")
    CloseTemp wb
End Sub

Public Sub Test_Ann_Stacked100_CagrUsesRawTotals()
    Dim wb As Workbook, cht As Chart, ln As Shape
    Set wb = NewTempWorkbook(): Set cht = KindFixture(wb, "stacked100")
    Set ln = Ann_AddCagr(cht, 1, 4)
    ' totais brutos: 2022 = 110+60 = 170; 2025 = 140+90 = 230 -> (230/170)^(1/3)-1 = 10,6%
    AssertEqual "CAGR +10,6%", cht.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text
    AssertAlmost 0, ln.Height, 0.5, "100%: ambos os topos em 1 -> linha plana"
    CloseTemp wb
End Sub

Public Sub Test_Ann_Stacked100_RefreshHidesExistingDiff()
    Dim wb As Workbook, cht As Chart, ln As Shape
    Set wb = NewTempWorkbook(): Set cht = KindFixture(wb, "stacked")
    Set ln = Ann_AddDifference(cht, 1, 4)
    AssertTrue ln.Visible = msoTrue
    cht.ChartType = xlColumnStacked100
    Registry_Block(cht).Cells(1, 2).Value = "stacked100"
    Ann_Refresh cht
    AssertTrue ln.Visible = msoFalse, "seta de diferenca oculta em 100%"
    AssertTrue cht.Shapes("tqq_ann_diff_1_lbl").Visible = msoFalse
    CloseTemp wb
End Sub

Public Sub Test_Ann_Refresh_IsolatesShapeFailure()
    Dim wb As Workbook, cht As Chart, a As Shape, b As Shape
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    Set a = Ann_AddCagr(cht, 1, 4)
    Set b = Ann_AddCagr(cht, 2, 3)
    ' sabota a primeira: troca o rotulo (caixa de texto) por uma linha com o mesmo nome -> TextRange falha
    cht.Shapes("tqq_ann_cagr_1_lbl").Delete
    cht.Shapes.AddLine(0, 0, 5, 5).Name = "tqq_ann_cagr_1_lbl"
    b.Left = 0: b.Top = 0
    Ann_Refresh cht
    AssertAlmost Geom_ChartCategoryCenterX(cht, 2), b.Left, 0.5, "segunda anotacao reposicionada apesar da falha da primeira"
    AssertTrue InStr(Ann_LastError(), "tqq_ann_cagr_1") > 0, "erro registrado: " & Ann_LastError()
    Ann_Refresh cht     ' a guarda foi liberada: nova chamada funciona
    CloseTemp wb
End Sub

Public Sub Test_Ann_Refresh_HidesStaleAndRestores()
    Dim wb As Workbook, cht As Chart, ln As Shape, ws As Worksheet, s As Series
    Set wb = NewTempWorkbook(): Set cht = KindFixture(wb, "stacked", 5)
    Set ws = wb.Worksheets("Dados")
    Set ln = Ann_AddCagr(cht, 1, 5)
    AssertTrue ln.Visible = msoTrue
    Set s = cht.SeriesCollection(1)
    Dim oldF As String: oldF = s.Formula
    s.Values = ws.Range("B2:B5")                 ' 4 pontos
    Ann_Refresh cht
    AssertTrue ln.Visible = msoFalse, "anotacao fora da faixa fica oculta"
    AssertTrue cht.Shapes("tqq_ann_cagr_1_lbl").Visible = msoFalse
    s.Formula = oldF
    Ann_Refresh cht
    AssertTrue ln.Visible = msoTrue, "volta a aparecer quando valida"
    AssertTrue cht.Shapes("tqq_ann_cagr_1_lbl").Visible = msoTrue
    CloseTemp wb
End Sub

Public Sub Test_Ann_Refresh_RebuildsWaterfallConnectors()
    Dim wb As Workbook, cht As Chart
    Set wb = NewTempWorkbook(): Set cht = WaterfallFixture(wb)
    AssertEqual 3, CountPrefix(cht, "tqq_ann_conn_")
    cht.Shapes("tqq_ann_conn_2").Delete
    Ann_Refresh cht
    AssertEqual 3, CountPrefix(cht, "tqq_ann_conn_")
    Registry_SetOption cht, "connectors", "0"
    cht.Shapes("tqq_ann_conn_2").Delete
    Ann_Refresh cht
    AssertEqual 2, CountPrefix(cht, "tqq_ann_conn_"), "conectores desligados nao sao recriados"
    CloseTemp wb
End Sub

Public Sub Test_Ann_Waterfall_Connectors_IndependentY()
    Dim wb As Workbook, cht As Chart, ax As Axis, it As Double, ih As Double
    Set wb = NewTempWorkbook(): Set cht = WaterfallFixture(wb)
    it = cht.PlotArea.InsideTop: ih = cht.PlotArea.InsideHeight: Set ax = cht.Axes(xlValue)
    ' End: 100, 70 (queda de 30), 120, 120
    AssertAlmost it + ih * (ax.MaximumScale - 70) / (ax.MaximumScale - ax.MinimumScale), cht.Shapes("tqq_ann_conn_2").Top, 1
    AssertAlmost it + ih * (ax.MaximumScale - 120) / (ax.MaximumScale - ax.MinimumScale), cht.Shapes("tqq_ann_conn_3").Top, 1
    CloseTemp wb
End Sub

Public Sub Test_Ann_Line_NotBetweenCategories()
    Dim wb As Workbook, cht As Chart, ln As Shape
    Set wb = NewTempWorkbook(): Set cht = KindFixture(wb, "line")
    cht.Axes(xlCategory).AxisBetweenCategories = False
    Set ln = Ann_AddCagr(cht, 1, 4)
    AssertAlmost cht.PlotArea.InsideLeft, ln.Left, 1
    AssertAlmost cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth, ln.Left + ln.Width, 1
    AssertAlmost 0, Geom_ChartBarHalfWidth(cht), 0.001, "linha nao tem barras"
    CloseTemp wb
End Sub

Public Sub Test_Ann_NewName_AvoidsCompanionCollisions()
    Dim wb As Workbook, cht As Chart, ln As Shape
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    cht.Shapes.AddLine(0, 0, 5, 5).Name = "tqq_ann_cagr_1_lbl"      ' orfao que colidiria com o rotulo
    Set ln = Ann_AddCagr(cht, 1, 4)
    AssertEqual "tqq_ann_cagr_2", ln.Name
    CloseTemp wb
End Sub

Public Sub Test_Ann_Labels_StayInsideChartArea()
    Dim wb As Workbook, cht As Chart, lbl As Shape
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    Ann_AddCagr cht, 1, 4
    Ann_AddDifference cht, 1, 4
    For Each lbl In cht.Shapes
        If Right$(lbl.Name, 4) = "_lbl" Then
            AssertTrue lbl.Left >= 0 And lbl.Top >= 0, lbl.Name & " nao pode sair por cima/esquerda"
            AssertTrue lbl.Left + lbl.Width <= cht.ChartArea.Width + 0.5, lbl.Name & " nao pode sair pela direita"
        End If
    Next
    CloseTemp wb
End Sub

Public Sub Test_Ann_CategoryData_EmptyAndShortSeries()
    Dim wb As Workbook, cht As Chart, v() As Double, t() As Double, ws As Worksheet
    Set wb = NewTempWorkbook(): Set cht = KindFixture(wb, "stacked", 4)
    Set ws = wb.Worksheets("Dados")
    cht.SeriesCollection(2).Values = ws.Range("C2:C3")     ' serie 2 mais curta que n: conta como zero
    Ann_GetCategoryData cht, v, t
    AssertEqual 4, UBound(v)
    AssertAlmost 140, v(4), 0.001, "serie curta conta como zero, sem erro"
    CloseTemp wb
End Sub

' Com os totais ligados, a linha do CAGR passa acima dos rotulos dos totais (posicao real do rotulo no grafico).
Public Sub Test_Ann_Cagr_ClearsTotalLabels_WhenTotalsOn()
    Dim wb As Workbook, cht As Chart, ln As Shape, s As Series, i As Long, lineYAtPoint As Double, labelTop As Double
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    AssertEqual "1", Registry_GetOption(cht, "totals", "0")
    Set ln = Ann_AddCagr(cht, 1, 4)
    Set s = Chart_FindSeries(cht, "tqq_total")
    ' linha sobe para a direita (VerticalFlip): inicio em ln.Top + ln.Height, fim em ln.Top
    lineYAtPoint = ln.Top + ln.Height
    labelTop = s.Points(1).DataLabel.Top
    AssertTrue lineYAtPoint <= labelTop - 8, "inicio da linha do CAGR acima do rotulo do total 1: linha=" & lineYAtPoint & " rotulo=" & labelTop
    lineYAtPoint = ln.Top
    labelTop = s.Points(4).DataLabel.Top
    AssertTrue lineYAtPoint <= labelTop - 8, "fim da linha do CAGR acima do rotulo do total 4: linha=" & lineYAtPoint & " rotulo=" & labelTop
    AssertTrue ln.Top >= 0, "linha dentro do grafico"
    CloseTemp wb
End Sub
' I1: categorias de data com intervalos irregulares: eixo de texto (uma fatia por categoria), nunca eixo de tempo.
Public Sub Test_Stacked_DateCategories_UseTextAxis_AndGeometryMatches()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, ln As Shape, expectEnd As Double
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Datas"
    ws.Range("A1:B1").Value = Array("Data", "Valor")
    ws.Range("A2").Value = DateSerial(2022, 1, 1): ws.Range("B2").Value = 100
    ws.Range("A3").Value = DateSerial(2023, 1, 1): ws.Range("B3").Value = 120
    ws.Range("A4").Value = DateSerial(2025, 1, 1): ws.Range("B4").Value = 200
    ws.Range("A2:A4").NumberFormat = "dd/mm/yyyy"
    Set d = ChartData_FromRange(ws.Range("A1:B4"))
    Set cht = Stacked_Insert(ws, d, "stacked").Chart
    AssertEqual xlCategoryScale, cht.Axes(xlCategory).CategoryType, "eixo de datas deve virar eixo de texto"
    Set ln = Ann_AddCagr(cht, 1, 3)
    ' regra de geometria independente: centro da 3a fatia de 3 = insideLeft + 2,5 * insideWidth / 3
    expectEnd = cht.PlotArea.InsideLeft + 2.5 * cht.PlotArea.InsideWidth / 3
    AssertAlmost expectEnd, ln.Left + ln.Width, 0.5, "fim da seta no centro da ultima categoria"
    CloseTemp wb
End Sub

' ---------- Tarefa 2: barras horizontais (fixture crua, sem depender da Tarefa 3) ----------

' Grafico de barras cru registrado como barstacked: bloco com categorias e totais, uma serie, estilo base.
Private Function RawBarFixture(wb As Workbook, Optional kind As String = "barstacked", Optional chartType As Long = xlBarStacked) As Chart
    Dim ws As Worksheet, blk As Range, body As Range, co As ChartObject, s As Series, i As Long
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:B1").Value = Array("Cat", "Valor")
    ws.Range("A2:B2").Value = Array("a", 10): ws.Range("A3:B3").Value = Array("b", 25): ws.Range("A4:B4").Value = Array("c", 40)
    Set blk = CalcSheet_AllocateBlock(wb, "geo001", kind, "'Dados'!$A$1:$B$4", 4, 4)
    Set body = Block_Body(blk)
    body.Cells(1, 1).Value = "Categoria": body.Cells(1, 2).Value = "Total"
    For i = 1 To 3
        body.Cells(i + 1, 1).Formula = "=Dados!$A$" & (i + 1)
        body.Cells(i + 1, 2).Formula = "=Dados!$B$" & (i + 1)
    Next
    Set co = ws.ChartObjects.Add(150, 10, 460, 280)
    co.Name = ID_PREFIX & "geo001"
    Set s = co.Chart.SeriesCollection.NewSeries
    s.Name = "=Dados!$B$1": s.Values = "=Dados!$B$2:$B$4": s.XValues = "=Dados!$A$2:$A$4"
    co.Chart.ChartType = chartType
    Style_ApplyBase co.Chart
    Registry_SetOption co.Chart, "numfmt", "int"
    Set RawBarFixture = co.Chart
End Function

' Borda livre a direita (valores reais do Excel): a da area do grafico ou, com legenda a direita, a borda esquerda dela.
Private Function RightObstacle(cht As Chart) As Double
    RightObstacle = cht.ChartArea.Width
    If cht.HasLegend Then
        If cht.Legend.Position = xlLegendPositionRight Then RightObstacle = cht.Legend.Left
    End If
End Function

' Rotulo do CAGR em barras: a direita da ponta da seta e dentro da borda livre.
Private Sub AssertCagrLabelClear(cht As Chart, ln As Shape, lbl As Shape, tag As String)
    AssertTrue lbl.Left >= ln.Left + ln.Width + 2, tag & ": rotulo a direita da ponta da seta (rotulo L=" & lbl.Left & ", ponta=" & (ln.Left + ln.Width) & ")"
    AssertTrue lbl.Left + lbl.Width <= RightObstacle(cht) + 0.5, tag & ": rotulo dentro da borda livre (direita do rotulo=" & (lbl.Left + lbl.Width) & ", borda=" & RightObstacle(cht) & ")"
End Sub

Public Sub Test_Geom_ChartCoords_HorizontalChart()
    Dim wb As Workbook, cht As Chart, ax As Axis, il As Double, it As Double, iw As Double, ih As Double
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    AssertTrue Geom_ChartIsHorizontal(cht)
    il = cht.PlotArea.InsideLeft: it = cht.PlotArea.InsideTop: iw = cht.PlotArea.InsideWidth: ih = cht.PlotArea.InsideHeight
    Set ax = cht.Axes(xlValue)
    AssertAlmost it + ih / 6, Geom_ChartCatCoord(cht, 1), 0.01, "categoria 1 no topo"
    AssertAlmost it + ih * 5 / 6, Geom_ChartCatCoord(cht, 3), 0.01
    AssertAlmost il + iw * (25 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale), Geom_ChartValCoord(cht, 25), 0.01
    AssertAlmost (ih / 3) / (1 + cht.ChartGroups(1).GapWidth / 100) / 2, Geom_ChartBarHalfWidth(cht), 0.01, "meia largura pela altura interna"
    CloseTemp wb
End Sub

Public Sub Test_Ann_Cagr_Horizontal_RunsAlongCategoriesRightOfBars()
    Dim wb As Workbook, cht As Chart, ln As Shape, lbl As Shape, ax As Axis, xEnd As Double
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    Set ln = Ann_AddCagr(cht, 1, 3)
    Set ax = cht.Axes(xlValue)
    xEnd = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (40 - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
    AssertAlmost Geom_ChartCatCoord(cht, 1), ln.Top, 1, "comeca na faixa da categoria 1"
    AssertAlmost Geom_ChartCatCoord(cht, 3) - Geom_ChartCatCoord(cht, 1), ln.Height, 1
    AssertAlmost xEnd + 10, ln.Left + ln.Width, 1.5, "fim da seta 10 pt a direita da barra mais longa"
    AssertTrue ln.VerticalFlip = msoFalse And ln.HorizontalFlip = msoFalse, "seta desce para a direita (ponta em 'c')"
    Set lbl = cht.Shapes("tqq_ann_cagr_1_lbl")
    AssertEqual "CAGR +100,0%", lbl.TextFrame2.TextRange.Text      ' (40/10)^(1/2)-1
    AssertTrue lbl.Left >= ln.Left + ln.Width + 2, "rotulo a direita da seta"
    AssertTrue lbl.Left + lbl.Width <= cht.ChartArea.Width + 0.5, "dentro do grafico"
    CloseTemp wb
End Sub

Public Sub Test_Ann_Diff_Horizontal_BelowLastBar()
    Dim wb As Workbook, cht As Chart, ln As Shape, lvl As Shape, lbl As Shape
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    Set ln = Ann_AddDifference(cht, 1, 3)
    Set lvl = cht.Shapes("tqq_ann_diff_1_lvl"): Set lbl = cht.Shapes("tqq_ann_diff_1_lbl")
    AssertAlmost 0, ln.Height, 0.01, "seta de diferenca horizontal"
    AssertTrue ln.Top >= Geom_ChartCatCoord(cht, 3) + Geom_ChartBarHalfWidth(cht) + 8, "abaixo da ultima barra"
    AssertAlmost Geom_ChartValCoord(cht, 10), ln.Left, 1
    AssertAlmost Geom_ChartValCoord(cht, 40) - Geom_ChartValCoord(cht, 10), ln.Width, 1
    AssertAlmost 0, lvl.Width, 0.01, "linha de nivel vertical"
    AssertAlmost Geom_ChartValCoord(cht, 10), lvl.Left, 1
    AssertEqual "+30 (+300%)", lbl.TextFrame2.TextRange.Text
    AssertTrue lbl.Top >= ln.Top + 4, "rotulo abaixo da seta"
    AssertTrue cht.ChartArea.Height - (cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight) >= 33.5, "folga abaixo da area de plotagem (Ann_EnsureRoom)"
    AssertTrue lbl.Top + lbl.Height <= cht.ChartArea.Height + 0.5, "rotulo da diferenca dentro do grafico"
    CloseTemp wb
End Sub

Public Sub Test_Ann_Refusals_ByKind()
    Dim wb As Workbook, cht As Chart, en As Long, ed As String, n0 As Long
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    n0 = cht.Shapes.Count
    Registry_Block(cht).Cells(1, 2).Value = "pie"            ' finge outro tipo: a recusa vem da matriz
    On Error Resume Next
    Ann_AddCagr cht, 1, 3
    en = Err.Number: ed = Err.Description: Err.Clear
    AssertEqual vbObjectError + 7001, en
    AssertEqual "A seta CAGR não está disponível para este tipo de gráfico.", ed
    AssertEqual n0, cht.Shapes.Count, "a recusa do CAGR nao cria forma"
    Ann_AddDifference cht, 1, 3
    en = Err.Number: ed = Err.Description: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7001, en
    AssertEqual "A seta de diferença não está disponível para este tipo de gráfico.", ed
    AssertEqual n0, cht.Shapes.Count, "a recusa da diferenca nao cria forma"
    Registry_Block(cht).Cells(1, 2).Value = "barstacked100"
    On Error Resume Next
    Ann_AddDifference cht, 1, 3
    en = Err.Number: ed = Err.Description: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7001, en
    AssertEqual "A seta de diferença não está disponível para gráficos 100%.", ed
    AssertEqual n0, cht.Shapes.Count, "a recusa em 100% nao cria forma"
    CloseTemp wb
End Sub

' O texto "gráficos 100%" e dos tipos 100% de colunas/barras/area; o Mekko (e demais) recebe a mensagem generica.
Public Sub Test_Ann_Refusals_DiffMessage_100PercentOnlyForColumnBarAreaKinds()
    Dim wb As Workbook, cht As Chart, ed As String, k As Variant, n0 As Long
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    n0 = cht.Shapes.Count
    For Each k In Array("stacked100", "barstacked100", "area100", "mekko", "mekkounits", "gantt", "pie", "doughnut")
        Registry_Block(cht).Cells(1, 2).Value = k
        On Error Resume Next
        Ann_AddDifference cht, 1, 3
        ed = Err.Description: Err.Clear
        On Error GoTo 0
        If k = "stacked100" Or k = "barstacked100" Or k = "area100" Then
            AssertEqual "A seta de diferença não está disponível para gráficos 100%.", ed, CStr(k)
        Else
            AssertEqual "A seta de diferença não está disponível para este tipo de gráfico.", ed, CStr(k)
        End If
    Next
    AssertEqual n0, cht.Shapes.Count, "nenhuma recusa cria forma"
    CloseTemp wb
End Sub

' Posicoes medidas nas barras desenhadas (Point.Left/Top/Width/Height), sem usar a formula de producao: o centro
' da faixa, o fim da barra, o zero e a meia espessura batem com Geom_Chart* (categoria 1 no topo, eixo cheio).
Public Sub Test_Geom_HorizontalCoords_MatchRenderedBars()
    Dim wb As Workbook, cht As Chart, pt As Point, i As Long, vals As Variant, ax As Axis
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    vals = Array(10, 25, 40)
    For i = 1 To 3
        Set pt = cht.SeriesCollection(1).Points(i)
        AssertAlmost pt.Top + pt.Height / 2, Geom_ChartCatCoord(cht, i), 0.6, "centro da faixa " & i
        AssertAlmost pt.Left + pt.Width, Geom_ChartValCoord(cht, CDbl(vals(i - 1))), 0.6, "fim da barra " & i
        AssertAlmost pt.Height / 2, Geom_ChartBarHalfWidth(cht), 0.6, "meia espessura " & i
    Next
    Set ax = cht.Axes(xlValue)
    ax.MinimumScale = -20: ax.MaximumScale = 50          ' escala fixa que inclui o negativo: o zero sai da borda
    For i = 1 To 3
        Set pt = cht.SeriesCollection(1).Points(i)
        AssertAlmost pt.Left, Geom_ChartValCoord(cht, 0), 0.6, "zero (inicio da barra " & i & ")"
        AssertAlmost pt.Left + pt.Width, Geom_ChartValCoord(cht, CDbl(vals(i - 1))), 0.6, "fim da barra " & i & " com escala fixa"
    Next
    CloseTemp wb
End Sub

' Conector horizontal: linha vertical no fim da barra i, da base da barra i ao topo da barra i+1.
Public Sub Test_Ann_Conn_Horizontal_VerticalBetweenBars()
    Dim wb As Workbook, cht As Chart, c As Shape, i As Long, a As Point, b As Point
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    Ann_SetWaterfallConnectors cht, True
    For i = 1 To 2
        Set c = cht.Shapes("tqq_ann_conn_" & i)
        Set a = cht.SeriesCollection(1).Points(i): Set b = cht.SeriesCollection(1).Points(i + 1)
        AssertAlmost 0, c.Width, 0.01, "conector " & i & " vertical"
        AssertAlmost a.Left + a.Width, c.Left, 0.6, "conector " & i & " no fim da barra " & i
        AssertAlmost a.Top + a.Height, c.Top, 0.6, "conector " & i & " sai da base da barra " & i
        AssertAlmost b.Top, c.Top + c.Height, 0.6, "conector " & i & " chega ao topo da barra " & (i + 1)
    Next
    CloseTemp wb
End Sub

' Ann_GetCategoryData por tipo de barra/area/combinacao: o tipo vem do bloco; as series sao as do grafico.
Public Sub Test_Ann_CategoryData_BarAreaComboKinds()
    Dim wb As Workbook, cht As Chart, ws As Worksheet, s As Series, v() As Double, t() As Double, k As Variant
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb): Set ws = wb.Worksheets("Dados")
    ws.Range("C1:C4").Value = Application.Transpose(Array("B", 5, -10, 20))
    ws.Range("D1:D4").Value = Application.Transpose(Array("L", 1, 2, 3))
    ws.Range("E1:E4").Value = Application.Transpose(Array("aux", 100, 100, 100))
    Set s = cht.SeriesCollection.NewSeries: s.Name = "=Dados!$C$1": s.Values = "=Dados!$C$2:$C$4": s.XValues = "=Dados!$A$2:$A$4"
    Set s = cht.SeriesCollection.NewSeries: s.Name = "=Dados!$D$1": s.Values = "=Dados!$D$2:$D$4": s.XValues = "=Dados!$A$2:$A$4"
    Set s = cht.SeriesCollection.NewSeries: s.Name = "tqq_total": s.Values = "=Dados!$E$2:$E$4": s.XValues = "=Dados!$A$2:$A$4"
    ' series: A = 10,25,40; B = 5,-10,20; L = 1,2,3; auxiliar tqq_total = 100 (nunca entra)
    For Each k In Array("barstacked", "area")
        Registry_Block(cht).Cells(1, 2).Value = k
        Ann_GetCategoryData cht, v, t
        AssertAlmost 16, v(1), 0.001, k & " soma 1": AssertAlmost 17, v(2), 0.001, k & " soma 2 (negativo entra)": AssertAlmost 63, v(3), 0.001, k & " soma 3"
        AssertAlmost 16, t(1), 0.001, k & " topo 1": AssertAlmost 27, t(2), 0.001, k & " topo 2 (so positivos)": AssertAlmost 63, t(3), 0.001, k & " topo 3"
    Next
    For Each k In Array("barstacked100", "area100")
        Registry_Block(cht).Cells(1, 2).Value = k
        Ann_GetCategoryData cht, v, t
        AssertAlmost 17, v(2), 0.001, k & " valores brutos": AssertAlmost 1, t(1), 0.001, k & " topo 1": AssertAlmost 1, t(2), 0.001, k & " topo 1": AssertAlmost 1, t(3), 0.001, k & " topo 1"
    Next
    Registry_Block(cht).Cells(1, 2).Value = "barclustered"
    Ann_GetCategoryData cht, v, t
    AssertAlmost 10, v(1), 0.001, "agrupada: serie 1": AssertAlmost 40, t(3), 0.001, "agrupada: serie 1"
    Registry_Block(cht).Cells(1, 2).Value = "combo"
    Ann_GetCategoryData cht, v, t
    AssertAlmost 15, v(1), 0.001, "combinacao: colunas A+B, sem a linha L": AssertAlmost 15, v(2), 0.001, "combinacao: colunas A+B": AssertAlmost 60, v(3), 0.001, "combinacao: colunas A+B"
    AssertAlmost 15, t(1), 0.001, "combinacao topo 1": AssertAlmost 25, t(2), 0.001, "combinacao topo 2 (so positivos)": AssertAlmost 60, t(3), 0.001, "combinacao topo 3"
    CloseTemp wb
End Sub

Public Sub Test_Ann_CategoryData_BarWaterfall_UsesEndColumn()
    Dim wb As Workbook, cht As Chart, v() As Double, t() As Double
    Set wb = NewTempWorkbook(): Set cht = WaterfallFixture(wb)
    Registry_Block(cht).Cells(1, 2).Value = "barwaterfall"
    Ann_GetCategoryData cht, v, t
    AssertAlmost 100, v(1), 0.001: AssertAlmost 70, v(2), 0.001: AssertAlmost 120, v(3), 0.001: AssertAlmost 120, t(4), 0.001
    CloseTemp wb
End Sub

' Em tipo que nao suporta seta, Ann_Refresh oculta a forma (ex.: CAGR em Mekko/pizza), em vez de reposiciona-la.
Public Sub Test_Ann_Refresh_HidesCagrWhenKindDoesNotSupport()
    Dim wb As Workbook, cht As Chart, ln As Shape
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    Set ln = Ann_AddCagr(cht, 1, 3)
    AssertTrue ln.Visible = msoTrue
    Registry_Block(cht).Cells(1, 2).Value = "pie"
    Ann_Refresh cht
    AssertTrue ln.Visible = msoFalse, "CAGR oculto em tipo sem suporte"
    AssertTrue cht.Shapes("tqq_ann_cagr_1_lbl").Visible = msoFalse
    Registry_Block(cht).Cells(1, 2).Value = "barstacked"
    Ann_Refresh cht
    AssertTrue ln.Visible = msoTrue, "volta quando o tipo suporta"
    CloseTemp wb
End Sub

' Ann_EnsureRoom em barras: 60 pt a direita da area interna com CAGR; so reserva embaixo quando ha seta de
' diferenca; e so aumenta a folga (nunca devolve largura a area de plotagem).
Public Sub Test_Ann_EnsureRoom_Horizontal_RightRoomAndOnlyGrows()
    Dim wb As Workbook, cht As Chart, ln As Shape, w0 As Double
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    Set ln = Ann_AddCagr(cht, 1, 3)
    AssertTrue RightObstacle(cht) - (cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth) >= 59.5, "folga a direita (ate a borda livre) para a seta e o rotulo do CAGR"
    AssertTrue cht.ChartArea.Height - (cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight) < 30, "sem seta de diferenca nada e reservado embaixo"
    cht.PlotArea.Width = cht.PlotArea.Width - 100          ' folga maior que a necessaria
    w0 = cht.PlotArea.InsideWidth
    Set ln = Ann_AddCagr(cht, 2, 3)
    AssertAlmost w0, cht.PlotArea.InsideWidth, 0.5, "so aumenta a folga: a largura nao e devolvida"
    Set ln = Ann_AddDifference(cht, 1, 3)
    AssertAlmost w0, cht.PlotArea.InsideWidth, 0.5, "a seta de diferenca so mexe na altura"
    AssertTrue cht.ChartArea.Height - (cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight) >= 33.5, "34 pt abaixo com a seta de diferenca"
    CloseTemp wb
End Sub

' Seta de diferenca sozinha nao reserva folga a direita (so o CAGR precisa dela).
Public Sub Test_Ann_EnsureRoom_Horizontal_DiffAloneKeepsRightSide()
    Dim wb As Workbook, cht As Chart, ln As Shape, r0 As Double
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    r0 = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth         ' borda direita da area interna (o Excel desloca a esquerda ao fixar a altura)
    Set ln = Ann_AddDifference(cht, 1, 3)
    AssertAlmost r0, cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth, 0.5, "diferenca sem CAGR nao move a borda direita da area de plotagem"
    AssertTrue cht.ChartArea.Height - (cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight) >= 33.5, "mas reserva 34 pt abaixo"
    CloseTemp wb
End Sub

' ---------- CAGR em barras: pior caso (a maior barra termina no fim do eixo) e legenda ----------

' (a) eixo com maximo fixo igual ao maior valor: a barra 3 (40) termina na borda direita da area interna.
Public Sub Test_Ann_Cagr_Horizontal_FixedMaxAtLongestBar_LabelClearOfArrow()
    Dim wb As Workbook, cht As Chart, ln As Shape, pt As Point
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    cht.Axes(xlValue).MinimumScale = 0: cht.Axes(xlValue).MaximumScale = 40
    Set ln = Ann_AddCagr(cht, 1, 3)
    Set pt = cht.SeriesCollection(1).Points(3)
    AssertAlmost cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth, pt.Left + pt.Width, 0.6, "premissa: a barra mais longa termina no fim do eixo"
    AssertAlmost pt.Left + pt.Width + 10, ln.Left + ln.Width, 1.5, "ponta da seta 10 pt depois da barra"
    AssertCagrLabelClear cht, ln, cht.Shapes("tqq_ann_cagr_1_lbl"), "maximo fixo"
    AssertTrue RightObstacle(cht) - (cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth) >= 79.5, "folga cheia (80 pt) quando a barra pode terminar no fim do eixo"
    CloseTemp wb
End Sub

' (a2) o mesmo sem legenda: a borda livre e a da area do grafico.
Public Sub Test_Ann_Cagr_Horizontal_FixedMaxAtLongestBar_NoLegend()
    Dim wb As Workbook, cht As Chart, ln As Shape
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    cht.HasLegend = False
    cht.Axes(xlValue).MinimumScale = 0: cht.Axes(xlValue).MaximumScale = 40
    Set ln = Ann_AddCagr(cht, 1, 3)
    AssertAlmost cht.ChartArea.Width, RightObstacle(cht), 0.001, "sem legenda a borda livre e a do grafico"
    AssertCagrLabelClear cht, ln, cht.Shapes("tqq_ann_cagr_1_lbl"), "sem legenda"
    CloseTemp wb
End Sub

' (b) com os totais ligados a seta se afasta 18 pt a mais; o rotulo continua livre.
Public Sub Test_Ann_Cagr_Horizontal_FixedMaxAtLongestBar_TotalsOn()
    Dim wb As Workbook, cht As Chart, ln As Shape, pt As Point
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    Registry_SetOption cht, "totals", "1"
    cht.Axes(xlValue).MinimumScale = 0: cht.Axes(xlValue).MaximumScale = 40
    Set ln = Ann_AddCagr(cht, 1, 3)
    Set pt = cht.SeriesCollection(1).Points(3)
    AssertAlmost pt.Left + pt.Width + 10 + 18, ln.Left + ln.Width, 1.5, "totais ligados: ponta 18 pt mais longe"
    AssertCagrLabelClear cht, ln, cht.Shapes("tqq_ann_cagr_1_lbl"), "maximo fixo com totais"
    AssertTrue RightObstacle(cht) - (cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth) >= 97.5, "folga cheia + 18 pt dos totais"
    CloseTemp wb
End Sub

' (a3) rotulo largo (CAGR de quatro digitos inteiros): a folga usa a largura medida do rotulo, nao um valor fixo.
Public Sub Test_Ann_Cagr_Horizontal_WideLabel_MeasuredWidthDrivesRoom()
    Dim wb As Workbook, cht As Chart, ln As Shape, lbl As Shape, ws As Worksheet
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb): Set ws = wb.Worksheets("Dados")
    ws.Range("B2").Value = 1: ws.Range("B3").Value = 10: ws.Range("B4").Value = 4000000       ' CAGR de 1 a 3 = +199.900,0%
    cht.Axes(xlValue).MinimumScale = 0: cht.Axes(xlValue).MaximumScale = 4000000
    Set ln = Ann_AddCagr(cht, 1, 3)
    Set lbl = cht.Shapes("tqq_ann_cagr_1_lbl")
    AssertTrue lbl.Width > 75, "premissa: rotulo largo (largura=" & lbl.Width & ")"
    AssertCagrLabelClear cht, ln, lbl, "rotulo largo"
    CloseTemp wb
End Sub

' (c) barras 100%: todas terminam no fim do eixo (100%). Fixture crua com tipo nativo xlBarStacked100 e bloco barstacked100
' (a Tarefa 3 pode trocar por Stacked_Insert(..., "barstacked100")).
Public Sub Test_Ann_Cagr_Horizontal_100Percent_LabelClearOfArrow()
    Dim wb As Workbook, cht As Chart, ln As Shape, pt As Point
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb, "barstacked100", xlBarStacked100)
    Set ln = Ann_AddCagr(cht, 1, 3)
    Set pt = cht.SeriesCollection(1).Points(3)
    AssertAlmost cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth, pt.Left + pt.Width, 0.6, "premissa: 100% termina no fim do eixo"
    AssertEqual "CAGR +100,0%", cht.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text      ' totais brutos 10 -> 40
    AssertCagrLabelClear cht, ln, cht.Shapes("tqq_ann_cagr_1_lbl"), "barras 100%"
    CloseTemp wb
End Sub

' Legenda a direita (o padrao de Stacked_Insert): a seta e o rotulo ficam a esquerda dela (valores reais da legenda).
Public Sub Test_Ann_Cagr_Horizontal_RightLegend_ArrowAndLabelLeftOfLegend()
    Dim wb As Workbook, cht As Chart, ln As Shape, lbl As Shape
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    AssertTrue cht.HasLegend And cht.Legend.Position = xlLegendPositionRight, "premissa: legenda a direita"
    Set ln = Ann_AddCagr(cht, 1, 3)
    Set lbl = cht.Shapes("tqq_ann_cagr_1_lbl")
    AssertTrue ln.Left + ln.Width <= cht.Legend.Left, "ponta da seta antes da legenda (ponta=" & (ln.Left + ln.Width) & ", legenda=" & cht.Legend.Left & ")"
    AssertTrue lbl.Left + lbl.Width <= cht.Legend.Left + 0.5, "rotulo antes da legenda (direita do rotulo=" & (lbl.Left + lbl.Width) & ", legenda=" & cht.Legend.Left & ")"
    AssertTrue lbl.Left >= ln.Left + ln.Width + 2, "rotulo a direita da ponta"
    AssertTrue cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth <= cht.Legend.Left - 59.5, "a area de plotagem deixa 60 pt ate a legenda"
    CloseTemp wb
End Sub

' Legenda embaixo com seta de diferenca: a seta e o rotulo ficam acima da legenda.
Public Sub Test_Ann_Diff_Horizontal_BottomLegend_ArrowAndLabelAboveLegend()
    Dim wb As Workbook, cht As Chart, ln As Shape, lbl As Shape
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb)
    cht.Legend.Position = xlLegendPositionBottom
    AssertTrue cht.Legend.Top > cht.ChartArea.Height / 2, "premissa: legenda na metade de baixo (topo=" & cht.Legend.Top & ")"
    Set ln = Ann_AddDifference(cht, 1, 3)
    Set lbl = cht.Shapes("tqq_ann_diff_1_lbl")
    AssertTrue ln.Top < cht.Legend.Top, "seta acima da legenda (seta=" & ln.Top & ", legenda=" & cht.Legend.Top & ")"
    AssertTrue lbl.Top + lbl.Height <= cht.Legend.Top + 0.5, "rotulo acima da legenda (base do rotulo=" & (lbl.Top + lbl.Height) & ", legenda=" & cht.Legend.Top & ")"
    AssertTrue lbl.Top >= ln.Top + 4, "rotulo abaixo da seta"
    AssertTrue cht.Legend.Top - (cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight) >= 33.5, "34 pt livres entre a area de plotagem e a legenda"
    CloseTemp wb
End Sub

' Cobertura da reconstrucao de conectores por familia (barwaterfall): o grafico de colunas serve so para contar formas;
' a cascata horizontal real chega na Tarefa 5.
Public Sub Test_Ann_Refresh_RebuildsConnectors_ForBarWaterfallFamily()
    Dim wb As Workbook, cht As Chart
    Set wb = NewTempWorkbook(): Set cht = WaterfallFixture(wb)
    Registry_Block(cht).Cells(1, 2).Value = "barwaterfall"
    AssertEqual 3, CountPrefix(cht, "tqq_ann_conn_")
    cht.Shapes("tqq_ann_conn_2").Delete
    Ann_Refresh cht
    AssertEqual 3, CountPrefix(cht, "tqq_ann_conn_"), "barwaterfall reconstroi os conectores faltantes"
    CloseTemp wb
End Sub

' Combinacao com menos de 2 series nao auxiliares: a unica serie conta (nao ha linha para excluir).
Public Sub Test_Ann_CategoryData_Combo_SingleSeries_IncludesIt()
    Dim wb As Workbook, cht As Chart, ws As Worksheet, s As Series, v() As Double, t() As Double
    Set wb = NewTempWorkbook(): Set cht = RawBarFixture(wb): Set ws = wb.Worksheets("Dados")
    Registry_Block(cht).Cells(1, 2).Value = "combo"
    Ann_GetCategoryData cht, v, t
    AssertAlmost 10, v(1), 0.001, "serie unica entra": AssertAlmost 40, v(3), 0.001, "serie unica entra": AssertAlmost 25, t(2), 0.001
    ws.Range("E1:E4").Value = Application.Transpose(Array("aux", 100, 100, 100))
    Set s = cht.SeriesCollection.NewSeries: s.Name = "tqq_total": s.Values = "=Dados!$E$2:$E$4": s.XValues = "=Dados!$A$2:$A$4"
    Ann_GetCategoryData cht, v, t
    AssertAlmost 10, v(1), 0.001, "auxiliar nao conta como serie de coluna nem como linha": AssertAlmost 40, v(3), 0.001
    CloseTemp wb
End Sub
