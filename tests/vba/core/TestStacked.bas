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

Private Function LocalizeFormat(fmt As String) As String
    Dim t As String, d As String
    t = Application.International(xlThousandsSeparator): d = Application.International(xlDecimalSeparator)
    LocalizeFormat = Replace(Replace(Replace(fmt, ",", "@@"), ".", d), "@@", t)
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
    AssertEqual LocalizeFormat("#,##0;-#,##0;"), cht.SeriesCollection(1).DataLabels.NumberFormat
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
