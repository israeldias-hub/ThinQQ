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
