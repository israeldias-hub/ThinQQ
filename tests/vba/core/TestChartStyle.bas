Attribute VB_Name = "TestChartStyle"
Option Explicit

Private Function MakeChart(wb As Workbook) As Chart
    Dim ws As Worksheet, co As ChartObject, s As Series
    Set ws = wb.Worksheets(1)
    ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Cat", "S1", "S2")
    ws.Range("A2:C2").Value = Array("a", 1, 2)
    ws.Range("A3:C3").Value = Array("b", 3, 4)
    Set co = ws.ChartObjects.Add(100, 10, 300, 200)
    Set s = co.Chart.SeriesCollection.NewSeries: s.Name = "=Dados!$B$1": s.Values = "=Dados!$B$2:$B$3"
    Set s = co.Chart.SeriesCollection.NewSeries: s.Name = "tqq_total": s.Values = "=Dados!$C$2:$C$3"
    co.Chart.ChartType = xlColumnStacked
    Set MakeChart = co.Chart
End Function

Public Sub Test_Style_Palettes_HaveSixColors()
    AssertEqual 6, UBound(Style_Palette("padrao")) - LBound(Style_Palette("padrao")) + 1
    AssertEqual 6, UBound(Style_Palette("cinza")) - LBound(Style_Palette("cinza")) + 1
    AssertEqual 6, UBound(Style_Palette("tema")) - LBound(Style_Palette("tema")) + 1
    AssertEqual RGB(31, 56, 100), Style_Palette("padrao")(0)
    AssertEqual RGB(31, 56, 100), Style_Palette("desconhecida")(0), "nome desconhecido cai no padrão"
End Sub

Public Sub Test_Style_NumberFormats()
    AssertEqual "#,##0;-#,##0;", Style_NumberFormat("int")
    AssertEqual "#,##0.0;-#,##0.0;", Style_NumberFormat("dec1")
    AssertEqual "0.0%;-0.0%;", Style_NumberFormat("pct1")
    AssertEqual "#,##0,;-#,##0,;", Style_NumberFormat("mil")
    AssertEqual "#,##0;-#,##0;", Style_NumberFormat("qualquer")
End Sub

Public Sub Test_Style_ContrastFont()
    AssertEqual RGB(255, 255, 255), Style_ContrastFont(RGB(31, 56, 100))
    AssertEqual RGB(64, 64, 64), Style_ContrastFont(RGB(217, 217, 217))
End Sub

Public Sub Test_Style_ApplyBase_HidesAxisAndGrid()
    Dim wb As Workbook, cht As Chart
    Set wb = NewTempWorkbook(): Set cht = MakeChart(wb)
    Style_ApplyBase cht
    AssertFalse cht.Axes(xlValue).HasMajorGridlines
    AssertEqual xlTickLabelPositionNone, cht.Axes(xlValue).TickLabelPosition
    AssertEqual msoFalse, cht.Axes(xlValue).Format.Line.Visible
    AssertEqual xlTickMarkNone, cht.Axes(xlCategory).MajorTickMark
    AssertFalse cht.HasTitle
    AssertAlmost 0, cht.Axes(xlValue).MinimumScale + 0   ' eixo ainda acessível (não removido)
    CloseTemp wb
End Sub

Public Sub Test_Style_SeriesColors_SkipHelpers()
    Dim wb As Workbook, cht As Chart
    Set wb = NewTempWorkbook(): Set cht = MakeChart(wb)
    Style_ApplySeriesColors cht, "cinza"
    AssertEqual RGB(64, 64, 64), cht.SeriesCollection(1).Format.Fill.ForeColor.RGB
    AssertTrue Style_IsHelperSeries(cht.SeriesCollection(2))
    AssertFalse Style_IsHelperSeries(cht.SeriesCollection(1))
    CloseTemp wb
End Sub

Public Sub Test_Style_ApplyNumberFormat_OnLabels()
    Dim wb As Workbook, cht As Chart
    Set wb = NewTempWorkbook(): Set cht = MakeChart(wb)
    cht.SeriesCollection(1).HasDataLabels = True
    Style_ApplyNumberFormat cht, Style_NumberFormat("dec2")

    AssertEqual "#,##0.00;-#,##0.00;", cht.SeriesCollection(1).DataLabels.NumberFormat
    CloseTemp wb
End Sub
