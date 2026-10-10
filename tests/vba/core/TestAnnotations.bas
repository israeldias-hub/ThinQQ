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
