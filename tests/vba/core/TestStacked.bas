Attribute VB_Name = "TestStacked"
Option Explicit

Private Function Fixture(wb As Workbook, Optional sheetName As String = "Dados") As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    ws.Name = sheetName
    ws.Range("A1:D1").Value = Array("Ano", "Sul", "Sudeste", "Nordeste")
    ws.Range("A2:D2").Value = Array(2022, 10, 20, 5)
    ws.Range("A3:D3").Value = Array(2023, 12, 22, 6)
    ws.Range("A4:D4").Value = Array(2024, 15, 25, 8)
    ws.Range("A5:D5").Value = Array(2025, 18, 30, 9)
    Set Fixture = ChartData_FromRange(ws.Range("A1:D5"))
End Function

Public Sub Test_Stacked_Insert_CreatesSeriesAndTotals()
    Dim wb As Workbook, d As clsChartData, co As ChartObject, cht As Chart, tot As Series, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set co = Stacked_Insert(d.Source.Worksheet, d, "stacked")
    Set cht = co.Chart
    AssertEqual ID_PREFIX, Left$(co.Name, 4)
    AssertEqual "stacked", Registry_ChartType(cht)
    AssertEqual 4, cht.SeriesCollection.Count
    AssertEqual "Sul", cht.SeriesCollection(1).Name
    AssertEqual xlColumnStacked, cht.SeriesCollection(1).ChartType
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertTrue Not tot Is Nothing, "série de totais ausente"
    AssertEqual xlLine, tot.ChartType
    AssertEqual msoFalse, tot.Format.Line.Visible
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual 35, body.Cells(2, 2).Value         ' 10+20+5
    AssertEqual 57, body.Cells(5, 2).Value         ' 18+30+9
    AssertTrue InStr(tot.Points(1).DataLabel.Formula, CALC_SHEET) > 0, "rótulo do total deve estar ligado ao bloco"
    AssertEqual xlLabelPositionAbove, tot.Points(1).DataLabel.Position
    AssertEqual "1", Registry_GetOption(cht, "totals")
    AssertEqual "1", Registry_GetOption(cht, "labels")
    AssertEqual "0", Registry_GetOption(cht, "connectors")
    AssertEqual "int", Registry_GetOption(cht, "numfmt")
    CloseTemp wb
End Sub

Public Sub Test_Stacked_SegmentLabels_CenteredWithFormat()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    AssertTrue cht.SeriesCollection(1).HasDataLabels
    AssertEqual xlLabelPositionCenter, cht.SeriesCollection(1).DataLabels.Position
    AssertEqual "#,##0;-#,##0;", cht.SeriesCollection(1).DataLabels.NumberFormat   ' getter devolve o código neutro (gravado por NumberFormatLocal)
    Stacked_SetLabels cht, False
    AssertFalse cht.SeriesCollection(1).HasDataLabels
    AssertEqual "0", Registry_GetOption(cht, "labels")
    CloseTemp wb
End Sub

Public Sub Test_Stacked_SetTotals_Off_RemovesSeries_On_Restores()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    Stacked_SetTotals cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertEqual 3, cht.SeriesCollection.Count
    Stacked_SetTotals cht, True
    AssertEqual 4, cht.SeriesCollection.Count
    CloseTemp wb
End Sub

Public Sub Test_Stacked_SetConnectors()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    Stacked_SetConnectors cht, True
    AssertTrue cht.ChartGroups(1).HasSeriesLines
    AssertEqual "1", Registry_GetOption(cht, "connectors")
    Stacked_SetConnectors cht, False
    AssertFalse cht.ChartGroups(1).HasSeriesLines
    CloseTemp wb
End Sub

Public Sub Test_Stacked100_TotalPosIsOne()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked100").Chart
    AssertEqual xlColumnStacked100, cht.SeriesCollection(1).ChartType
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual 1, body.Cells(2, 3).Value
    AssertEqual 35, body.Cells(2, 2).Value
    AssertTrue Not Chart_FindSeries(cht, "tqq_total") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_Clustered_LabelsOutsideEnd_NoTotals()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "clustered").Chart
    AssertEqual xlColumnClustered, cht.SeriesCollection(1).ChartType
    AssertEqual xlLabelPositionOutsideEnd, cht.SeriesCollection(1).DataLabels.Position
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertEqual "80", Registry_GetOption(cht, "gap")
    CloseTemp wb
End Sub

Public Sub Test_Line_Insert()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "line").Chart
    AssertEqual xlLineMarkers, cht.SeriesCollection(1).ChartType
    AssertEqual xlLabelPositionAbove, cht.SeriesCollection(1).DataLabels.Position
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_Stacked_TextCellTreatedAsZero()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Worksheet.Range("C3").Value = "n/d"
    d.Source.Worksheet.Range("D4").ClearContents
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual 18, body.Cells(3, 2).Value       ' 12 + "n/d" + 6
    AssertEqual 40, body.Cells(4, 2).Value       ' 15 + 25 + vazio
    CloseTemp wb
End Sub

Public Sub Test_Stacked_NegativeTotal_LabelBelow()
    Dim wb As Workbook, d As clsChartData, cht As Chart, tot As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Worksheet.Range("B2:D2").Value = Array(-10, -20, -5)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertEqual xlLabelPositionBelow, tot.Points(1).DataLabel.Position
    AssertEqual xlLabelPositionAbove, tot.Points(2).DataLabel.Position
    d.Source.Worksheet.Range("B2:D2").Value = Array(10, 20, 5)
    Labels_RefreshPositions cht
    AssertEqual xlLabelPositionAbove, tot.Points(1).DataLabel.Position
    CloseTemp wb
End Sub

Public Sub Test_Stacked_Insert_SheetNameWithQuote()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb, "Plano 'A' 2026")
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    AssertEqual "Sul", cht.SeriesCollection(1).Name
    AssertTrue InStr(cht.SeriesCollection(1).Formula, "''A''") > 0, "apóstrofos dobrados na fórmula da série"
    AssertTrue InStr(Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Formula, CALC_SHEET) > 0, "rótulo ligado ao bloco"
    AssertEqual 35, Block_Body(Registry_Block(cht)).Cells(2, 2).Value
    CloseTemp wb
End Sub

Public Sub Test_Stacked_Legend_DoesNotListTotal()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    AssertEqual 3, cht.Legend.LegendEntries.Count
    Stacked_SetTotals cht, False
    Stacked_SetTotals cht, True
    AssertEqual 3, cht.Legend.LegendEntries.Count, "legenda após reativar totais"
    CloseTemp wb
End Sub

Public Sub Test_Stacked_Legend_ExtraSeries_KeepsDataEntries()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, n1 As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    AssertEqual 3, cht.Legend.LegendEntries.Count, "só as 3 séries de dados"
    Set s = cht.SeriesCollection.NewSeries
    s.Name = "Extra": s.Values = "=Dados!$B$2:$B$5"
    n1 = cht.Legend.LegendEntries.Count
    Stacked_HideTotalLegendEntry cht
    Stacked_HideTotalLegendEntry cht
    AssertEqual n1, cht.Legend.LegendEntries.Count, "chamada repetida não remove entrada real"
    AssertTrue n1 >= 3, "entradas das séries de dados preservadas"
    ' religar a legenda: a entrada do total volta e deve ser removida de novo, sem tocar nas reais
    cht.HasLegend = False: cht.HasLegend = True
    AssertEqual cht.SeriesCollection.Count, cht.Legend.LegendEntries.Count
    Stacked_HideTotalLegendEntry cht
    AssertEqual cht.SeriesCollection.Count - 1, cht.Legend.LegendEntries.Count
    CloseTemp wb
End Sub

Public Sub Test_Stacked_ErrorCell_IgnoredInTotal()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range, i As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Worksheet.Range("C2").Formula = "=NA()"
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual 15, body.Cells(2, 2).Value
    For i = 2 To 5
        AssertFalse IsError(body.Cells(i, 2).Value), "total com erro"
        AssertFalse IsError(body.Cells(i, 3).Value), "totalpos com erro"
    Next
    CloseTemp wb
End Sub

Public Sub Test_Stacked_SetTotals_OnlyForStackedKinds()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "clustered").Chart
    Stacked_SetTotals cht, True
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertEqual 3, cht.SeriesCollection.Count
    Stacked_SetTotals cht, False
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "line").Chart
    Stacked_SetTotals cht, True
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertEqual "0", Registry_GetOption(cht, "totals")
    CloseTemp wb
End Sub

Public Sub Test_Stacked_TwoInserts_DoNotOverlap()
    Dim wb As Workbook, d As clsChartData, a As ChartObject, b As ChartObject
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set a = Stacked_Insert(d.Source.Worksheet, d, "stacked")
    Set b = Stacked_Insert(d.Source.Worksheet, d, "stacked")
    AssertTrue b.Top >= a.Top + a.Height, "segundo gráfico sobrepõe o primeiro"
    CloseTemp wb
End Sub

Public Sub Test_Stacked_Connectors_Line_PersistsZero()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "line").Chart
    Stacked_SetConnectors cht, True
    AssertEqual "0", Registry_GetOption(cht, "connectors")
    CloseTemp wb
End Sub

Public Sub Test_Stacked_LabelContrast_FollowsFill()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, i As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:G1").Value = Array("Ano", "S1", "S2", "S3", "S4", "S5", "S6")
    ws.Range("A2:G2").Value = Array(2022, 1, 2, 3, 4, 5, 6)
    ws.Range("A3:G3").Value = Array(2023, 1, 2, 3, 4, 5, 6)
    Set d = ChartData_FromRange(ws.Range("A1:G3"))
    Set cht = Stacked_Insert(ws, d, "stacked").Chart
    AssertEqual RGB(255, 255, 255), cht.SeriesCollection(1).DataLabels.Font.Color
    Style_ApplySeriesColors cht, "cinza"
    AssertEqual RGB(255, 255, 255), cht.SeriesCollection(1).DataLabels.Font.Color
    AssertEqual RGB(64, 64, 64), cht.SeriesCollection(5).DataLabels.Font.Color
    CloseTemp wb
End Sub
