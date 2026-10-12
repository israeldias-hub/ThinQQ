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

' Onda A: 10 cores por paleta (pizza e rosca de 7 a 10 fatias sem repetir); as 6 primeiras sao as de antes (series e pizzas de ate 6
' fatias nao mudam).
Public Sub Test_Style_Palettes_HaveTenColors_FirstSixUnchanged()
    Dim p As Variant, i As Long, j As Long, k As Variant
    AssertEqual 10, UBound(Style_Palette("padrao")) - LBound(Style_Palette("padrao")) + 1
    AssertEqual 10, UBound(Style_Palette("cinza")) - LBound(Style_Palette("cinza")) + 1
    AssertEqual 10, UBound(Style_Palette("tema")) - LBound(Style_Palette("tema")) + 1
    p = Array(RGB(31, 56, 100), RGB(68, 114, 196), RGB(157, 195, 230), RGB(165, 165, 165), RGB(0, 150, 136), RGB(237, 162, 36))
    For i = 0 To 5: AssertEqual CLng(p(i)), CLng(Style_Palette("padrao")(i)), "padrao " & i & " como antes": Next
    p = Array(RGB(64, 64, 64), RGB(127, 127, 127), RGB(166, 166, 166), RGB(191, 191, 191), RGB(217, 217, 217), RGB(89, 89, 89))
    For i = 0 To 5: AssertEqual CLng(p(i)), CLng(Style_Palette("cinza")(i)), "cinza " & i & " como antes": Next
    For i = 0 To 5: AssertEqual -(i + 1), CLng(Style_Palette("tema")(i)), "tema " & i & " como antes": Next
    For Each k In Array("padrao", "cinza", "tema")
        p = Style_Palette(CStr(k))
        For i = 0 To 9
            For j = i + 1 To 9
                AssertTrue p(i) <> p(j), k & ": cores " & i & " e " & j & " iguais"
            Next
        Next
    Next
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

' Onda A, rodada 1 (menor d): a secao limpa do formato tira textos, colchetes ([Red] tem um "d", que parecia data) e literais.
Public Sub Test_Style_CleanFormatSection()
    AssertEqual "0.0%", Style_CleanFormatSection("[Red]0.0%;-0.0%")
    AssertEqual " #,##0.00", Style_CleanFormatSection("""R$"" #,##0.00")
    AssertEqual "#,##0", Style_CleanFormatSection("[$R$-416]#,##0")
    AssertEqual "0", Style_CleanFormatSection("\d0")
    AssertEqual "dd/mm/yyyy", Style_CleanFormatSection("dd/mm/yyyy")
End Sub
