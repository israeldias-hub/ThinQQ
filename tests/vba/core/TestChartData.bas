Attribute VB_Name = "TestChartData"
Option Explicit

Private Function Fixture(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Ano", "Sul", "Sudeste", "Nordeste")
    ws.Range("A2:D2").Value = Array(2022, 10, 20, 5)
    ws.Range("A3:D3").Value = Array(2023, 12, 22, 6)
    ws.Range("A4:D4").Value = Array(2024, 15, 25, 8)
    ws.Range("A5:D5").Value = Array(2025, 18, 30, 9)
    Set Fixture = ws
End Function

Public Sub Test_ChartData_Shape()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set d = ChartData_FromRange(ws.Range("A1:D5"))
    AssertEqual 4, d.CategoryCount
    AssertEqual 3, d.SeriesCount
    AssertEqual "Sudeste", d.SeriesName(2)
    AssertEqual "2023", d.CategoryText(2)
    AssertEqual "$A$2:$A$5", d.CategoryCells().Address
    AssertEqual "$C$2:$C$5", d.SeriesValueCells(2).Address
    AssertEqual "$C$1", d.SeriesNameCell(2).Address
    AssertEqual 25, d.ValueCell(3, 2).Value
    AssertEqual "", d.Validate()
    CloseTemp wb
End Sub

Public Sub Test_ChartData_SingleCell_ExpandsToCurrentRegion()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set d = ChartData_FromRange(ws.Range("B3"))
    AssertEqual "$A$1:$D$5", d.Source.Address
    CloseTemp wb
End Sub

Public Sub Test_ChartData_ExternalAddress_EscapesQuote()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Name = "Plano 'A' 2026"
    AssertEqual "='Plano ''A'' 2026'!$B$2:$B$5", ChartData_ExternalAddress(ws.Range("B2:B5"))
    CloseTemp wb
End Sub

Public Sub Test_ChartData_Validate_TooSmall()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set d = ChartData_FromRange(ws.Range("A1:A5"))
    AssertTrue Len(d.Validate()) > 0, "uma coluna só deveria ser inválido"
    Set d = ChartData_FromRange(ws.Range("A1:D1"))
    AssertTrue Len(d.Validate()) > 0, "uma linha só deveria ser inválido"
    CloseTemp wb
End Sub

Public Sub Test_ChartData_Validate_MultiArea()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set d = ChartData_FromRange(ws.Range("A1:B5,D1:D5"))
    AssertTrue Len(d.Validate()) > 0
    CloseTemp wb
End Sub

Public Sub Test_ChartData_Validate_SingleEmptyCell()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set d = ChartData_FromRange(ws.Range("H20"))
    AssertTrue InStr(d.Validate(), "mínimo 2 linhas x 2 colunas") > 0, d.Validate()
    CloseTemp wb
End Sub

Public Sub Test_ChartData_FullColumns_LimitedToUsedRange()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set d = ChartData_FromRange(ws.Range("A:D"))
    AssertEqual 4, d.CategoryCount
    AssertEqual 3, d.SeriesCount
    AssertEqual "", d.Validate()
    Set d = ChartData_FromRange(ws.Rows("1:5"))
    AssertEqual 4, d.CategoryCount
    CloseTemp wb
End Sub

Public Sub Test_ChartData_Validate_TooLarge()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set d = New clsChartData
    d.Init ws.Range("A1:D100005")
    AssertTrue InStr(d.Validate(), "muito grande") > 0, d.Validate()
    CloseTemp wb
End Sub

Public Sub Test_ChartData_Validate_CalcSheetRejected()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    wb.Worksheets.Add.Name = CALC_SHEET
    Set ws = wb.Worksheets(CALC_SHEET)
    ws.Range("A1:B2").Value = 1
    Set d = ChartData_FromRange(ws.Range("A1:B2"))
    AssertTrue InStr(d.Validate(), "planilha interna") > 0, d.Validate()
    CloseTemp wb
End Sub

Public Sub Test_ChartData_Validate_MultiArea_Message()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set d = ChartData_FromRange(ws.Range("A1:B5,D1:D5"))
    AssertTrue InStr(d.Validate(), "única área") > 0, d.Validate()
    CloseTemp wb
End Sub

Public Sub Test_ChartData_TextAndBlankCells_Accessible()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, v As Variant
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("C3").Value = "n/d"
    ws.Range("D4").ClearContents
    Set d = ChartData_FromRange(ws.Range("A1:D5"))
    AssertEqual 4, d.CategoryCount
    AssertEqual 3, d.SeriesCount
    AssertEqual "", d.Validate()
    AssertEqual "n/d", d.ValueCell(2, 2).Value
    v = d.ValueCell(3, 3).Value
    AssertTrue IsEmpty(v), "célula vazia"
    CloseTemp wb
End Sub

Public Sub Test_ChartData_Text_NarrowColumn()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A3").Value = 2023
    ws.Range("A3").NumberFormat = "#,##0.00000"
    ws.Range("B1").Value = "Um cabeçalho bastante longo"
    ws.Columns("A:B").ColumnWidth = 1
    Set d = ChartData_FromRange(ws.Range("A1:D5"))
    AssertEqual "2023", d.CategoryText(2)
    AssertEqual "Um cabeçalho bastante longo", d.SeriesName(1)
    CloseTemp wb
End Sub

Public Sub Test_Runner_ClosesLeakedWorkbook()
    Dim wb As Workbook
    Set wb = NewTempWorkbook()
    AssertTrue wb.Worksheets.Count > 0   ' o runner deve fechar este workbook ao fim do teste
End Sub

Public Sub Test_Runner_LeakedWorkbookWasClosed()
    Dim w As Workbook, n As Long
    For Each w In Application.Workbooks
        If Not w.IsAddin Then n = n + 1
    Next
    AssertEqual 0, n, "pastas vazadas permaneceram abertas"
End Sub
