Attribute VB_Name = "TestWaterfall"
Option Explicit

Private Function LocalizeFormat(fmt As String) As String
    Dim t As String, d As String
    t = Application.International(xlThousandsSeparator): d = Application.International(xlDecimalSeparator)
    LocalizeFormat = Replace(Replace(Replace(fmt, ",", "@@"), ".", d), "@@", t)
End Function

Private Function Fixture(wb As Workbook, Optional sheetName As String = "Ponte") As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    ws.Name = sheetName
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Início", 100)
    ws.Range("A3:B3").Value = Array("A", -150)
    ws.Range("A4:B4").Value = Array("B", 30)
    ws.Range("A5:B5").Value = Array("Subtotal", "e")
    ws.Range("A6:B6").Value = Array("C", 70)
    ws.Range("A7:B7").Value = Array("Fim", "E")
    Set Fixture = ChartData_FromRange(ws.Range("A1:B7"))
End Function

Public Sub Test_Waterfall_IsTotalMarker()
    AssertTrue Waterfall_IsTotalMarker("e")
    AssertTrue Waterfall_IsTotalMarker("E")
    AssertTrue Waterfall_IsTotalMarker(" total ")
    AssertTrue Waterfall_IsTotalMarker("=")
    AssertFalse Waterfall_IsTotalMarker(12)
    AssertFalse Waterfall_IsTotalMarker("")
    AssertFalse Waterfall_IsTotalMarker("abc")
End Sub

Public Sub Test_Waterfall_Block_CrossingZero()
    Dim wb As Workbook, d As clsChartData, body As Range, r As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wf0001", d))
    ' End: 100, -50, -20, -20, 50, 50
    AssertEqual 100, body.Cells(2, WF_COL_END).Value
    AssertEqual -50, body.Cells(3, WF_COL_END).Value
    AssertEqual -20, body.Cells(4, WF_COL_END).Value
    AssertEqual -20, body.Cells(5, WF_COL_END).Value
    AssertEqual 50, body.Cells(6, WF_COL_END).Value
    AssertEqual 50, body.Cells(7, WF_COL_END).Value
    ' Linha A (100 -> -50, desce cruzando o zero): Base 0, DownPos 100, DownNeg -50
    r = 3
    AssertEqual 0, body.Cells(r, WF_COL_BASE).Value
    AssertEqual 100, body.Cells(r, 14).Value
    AssertEqual -50, body.Cells(r, 15).Value
    AssertEqual 0, body.Cells(r, 12).Value
    ' Linha B (-50 -> -20, sobe no negativo): Base -20, UpNeg -30
    r = 4
    AssertEqual -20, body.Cells(r, WF_COL_BASE).Value
    AssertEqual -30, body.Cells(r, 13).Value
    AssertEqual 0, body.Cells(r, 12).Value
    ' Subtotal (e): Total -20, rótulo -20 posicionado em Lo (-20)
    r = 5
    AssertEqual True, body.Cells(r, 3).Value
    AssertEqual -20, body.Cells(r, 16).Value
    AssertEqual -20, body.Cells(r, WF_COL_LABELVAL).Value
    AssertEqual -20, body.Cells(r, WF_COL_LABELPOS).Value
    ' Linha C (-20 -> 50, sobe cruzando o zero): Base 0, UpPos 50, UpNeg -20; rótulo 70 em Hi (50)
    r = 6
    AssertEqual 0, body.Cells(r, WF_COL_BASE).Value
    AssertEqual 50, body.Cells(r, 12).Value
    AssertEqual -20, body.Cells(r, 13).Value
    AssertEqual 70, body.Cells(r, WF_COL_LABELVAL).Value
    AssertEqual 50, body.Cells(r, WF_COL_LABELPOS).Value
    ' Início: Base 0, UpPos 100
    AssertEqual 0, body.Cells(2, WF_COL_BASE).Value
    AssertEqual 100, body.Cells(2, 12).Value
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_Block_LiveUpdate()
    Dim wb As Workbook, d As clsChartData, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wf0002", d))
    d.Source.Worksheet.Range("B3").Value = -10
    AssertEqual 190, body.Cells(7, WF_COL_END).Value     ' 100-10+30+70
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_Insert_Series()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lbl As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "waterfall", Registry_ChartType(cht)
    AssertEqual 7, cht.SeriesCollection.Count
    AssertEqual "tqq_base", cht.SeriesCollection(1).Name
    AssertEqual xlColumnStacked, cht.SeriesCollection(1).ChartType
    AssertEqual msoFalse, cht.SeriesCollection(1).Format.Fill.Visible
    AssertFalse cht.HasLegend
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    AssertEqual xlLine, lbl.ChartType
    AssertTrue InStr(lbl.Points(3).DataLabel.Formula, CALC_SHEET) > 0
    AssertEqual xlLabelPositionBelow, lbl.Points(4).DataLabel.Position   ' subtotal -20
    AssertEqual xlLabelPositionAbove, lbl.Points(5).DataLabel.Position   ' +70
    AssertEqual "1", Registry_GetOption(cht, "labels")
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_EndValues_And_LabelsOff()
    Dim wb As Workbook, d As clsChartData, cht As Chart, v() As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    v = Waterfall_EndValues(cht)
    AssertEqual 6, UBound(v)
    AssertAlmost -50, v(2)
    AssertAlmost 50, v(6)
    Waterfall_SetLabels cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl") Is Nothing
    AssertEqual 6, cht.SeriesCollection.Count
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_Block_MarkerEdgeCases()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, body As Range
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("I", 10)
    ws.Range("A3").Value = "F": ws.Range("B3").Formula = "="""""    ' formula devolvendo vazio: nao e total
    ws.Range("A4:B4").Value = Array("G", " Total ")
    ws.Range("A5:B5").Value = Array("H", "n/d")
    ws.Range("A6").Value = "J": ws.Range("B6").NumberFormat = "@": ws.Range("B6").Value = "100"
    Set d = ChartData_FromRange(ws.Range("A1:B6"))
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wf0003", d))
    AssertEqual False, body.Cells(3, 3).Value
    AssertEqual True, body.Cells(4, 3).Value
    AssertEqual False, body.Cells(5, 3).Value
    AssertEqual False, body.Cells(6, 3).Value
    AssertEqual 10, body.Cells(4, WF_COL_END).Value     ' total: delta 0, end = start
    AssertEqual 10, body.Cells(6, WF_COL_END).Value     ' numero como texto conta 0
    CloseTemp wb
End Sub

' Monta uma tabela simples (A: categoria, B: valor) a partir de um array de valores.
Private Function Simple(wb As Workbook, vals As Variant) As clsChartData
    Dim ws As Worksheet, i As Long
    Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    For i = 0 To UBound(vals)
        ws.Cells(i + 2, 1).Value = "C" & i
        ws.Cells(i + 2, 2).Value = vals(i)
    Next
    Set Simple = ChartData_FromRange(ws.Range(ws.Cells(1, 1), ws.Cells(UBound(vals) + 2, 2)))
End Function

Public Sub Test_Waterfall_Block_LoHiPieces()
    Dim wb As Workbook, body As Range
    Set wb = NewTempWorkbook()
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wf0010", Fixture(wb)))
    ' Lo=7 Hi=8 PiecePos=10 PieceNeg=11
    AssertEqual 0, body.Cells(2, 7).Value: AssertEqual 100, body.Cells(2, 8).Value
    AssertEqual 100, body.Cells(2, 10).Value: AssertEqual 0, body.Cells(2, 11).Value
    AssertEqual -50, body.Cells(3, 7).Value: AssertEqual 100, body.Cells(3, 8).Value
    AssertEqual 100, body.Cells(3, 10).Value: AssertEqual -50, body.Cells(3, 11).Value
    AssertEqual -50, body.Cells(4, 7).Value: AssertEqual -20, body.Cells(4, 8).Value
    AssertEqual 0, body.Cells(4, 10).Value: AssertEqual -30, body.Cells(4, 11).Value
    AssertEqual -20, body.Cells(5, 7).Value: AssertEqual 0, body.Cells(5, 8).Value
    AssertEqual 0, body.Cells(5, 10).Value: AssertEqual -20, body.Cells(5, 11).Value
    AssertEqual -20, body.Cells(6, 7).Value: AssertEqual 50, body.Cells(6, 8).Value
    AssertEqual 50, body.Cells(6, 10).Value: AssertEqual -20, body.Cells(6, 11).Value
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_Block_AllNegative()
    Dim wb As Workbook, body As Range
    Set wb = NewTempWorkbook()
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wf0011", Simple(wb, Array(-10, -20, 5))))
    AssertEqual -10, body.Cells(2, WF_COL_END).Value
    AssertEqual -30, body.Cells(3, WF_COL_END).Value
    AssertEqual -25, body.Cells(4, WF_COL_END).Value
    AssertEqual 0, body.Cells(2, WF_COL_BASE).Value
    AssertEqual -10, body.Cells(2, 15).Value
    AssertEqual -10, body.Cells(3, WF_COL_BASE).Value
    AssertEqual -20, body.Cells(3, 15).Value
    AssertEqual 0, body.Cells(3, 14).Value
    AssertEqual -25, body.Cells(4, WF_COL_BASE).Value
    AssertEqual -5, body.Cells(4, 13).Value
    AssertEqual -25, body.Cells(4, WF_COL_LABELPOS).Value   ' +5 em Hi (-25)
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_Block_ZeroDelta()
    Dim wb As Workbook, body As Range
    Set wb = NewTempWorkbook()
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wf0012", Simple(wb, Array(10, 0, "abc"))))
    AssertEqual 10, body.Cells(3, WF_COL_END).Value
    AssertEqual 10, body.Cells(3, WF_COL_BASE).Value
    AssertEqual 0, body.Cells(3, 12).Value
    AssertEqual 0, body.Cells(3, 14).Value
    AssertEqual 0, body.Cells(3, WF_COL_LABELVAL).Value     ' formato "#,##0;-#,##0;" exibe zero vazio
    AssertEqual 10, body.Cells(4, WF_COL_END).Value         ' texto nao marcador = delta 0
    AssertEqual 0, body.Cells(4, WF_COL_LABELVAL).Value
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_Block_TotalFirstRow()
    Dim wb As Workbook, body As Range, c As Long
    Set wb = NewTempWorkbook()
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wf0013", Simple(wb, Array("e", 5))))
    AssertEqual 0, body.Cells(2, WF_COL_END).Value
    For c = 9 To 16
        AssertEqual 0, body.Cells(2, c).Value, "coluna " & c
    Next
    AssertEqual 5, body.Cells(3, WF_COL_END).Value
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_Block_ErrorCellTreatedAsZero()
    Dim wb As Workbook, d As clsChartData, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Worksheet.Range("B3").Formula = "=NA()"
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wf0014", d))
    AssertEqual 100, body.Cells(3, WF_COL_END).Value
    AssertEqual False, body.Cells(3, 3).Value
    AssertEqual 130, body.Cells(4, WF_COL_END).Value
    AssertEqual 130, body.Cells(5, WF_COL_END).Value
    AssertEqual 200, body.Cells(6, WF_COL_END).Value
    AssertEqual 200, body.Cells(7, WF_COL_END).Value
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_MarkerList_FormulaMatchesFunction()
    Dim wb As Workbook, d As clsChartData, body As Range, m As Variant, i As Long
    m = Array("e", "t", "=", "total", "subtotal", "E", "TOTAL", " e ", "x", "tt", "")
    Set wb = NewTempWorkbook()
    Set d = Simple(wb, Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1))
    For i = 0 To UBound(m)
        d.Source.Worksheet.Cells(i + 2, 2).NumberFormat = "@"
        d.Source.Worksheet.Cells(i + 2, 2).Value = CStr(m(i))
    Next
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wf0015", d))
    For i = 0 To UBound(m)
        AssertEqual Waterfall_IsTotalMarker(CStr(m(i))), CBool(body.Cells(i + 2, 3).Value), "marcador '" & m(i) & "'"
    Next
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_Insert_SheetNameWithQuote()
    Dim wb As Workbook, d As clsChartData, cht As Chart, v As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb, "Plano 'A' 2026")
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    v = cht.SeriesCollection(2).Values
    AssertEqual 100, v(1)
    AssertTrue InStr(cht.SeriesCollection(1).Formula, CALC_SHEET) > 0
    AssertTrue InStr(cht.SeriesCollection(1).Formula, "''A''") > 0, "categorias na planilha com apostrofo"
    AssertTrue InStr(Chart_FindSeries(cht, "tqq_lbl").Points(2).DataLabel.Formula, CALC_SHEET) > 0
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_TwoInserts_DoNotOverlap()
    Dim wb As Workbook, d As clsChartData, c1 As ChartObject, c2 As ChartObject
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set c1 = Waterfall_Insert(d.Source.Worksheet, d)
    Set c2 = Waterfall_Insert(d.Source.Worksheet, d)
    AssertTrue c2.Top >= c1.Top + c1.Height, "segundo grafico abaixo do primeiro"
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_Insert_ColorsAndLabelFormat()
    Dim wb As Workbook, d As clsChartData, cht As Chart, pal As Variant, lbl As Series, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    pal = Style_Palette("padrao")
    AssertEqual msoFalse, Chart_FindSeries(cht, "tqq_base").Format.Fill.Visible
    AssertEqual CLng(pal(1)), Chart_FindSeries(cht, "tqq_uppos").Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(1)), Chart_FindSeries(cht, "tqq_upneg").Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(2)), Chart_FindSeries(cht, "tqq_downpos").Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(2)), Chart_FindSeries(cht, "tqq_downneg").Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(0)), Chart_FindSeries(cht, "tqq_total").Format.Fill.ForeColor.RGB
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual -150, body.Cells(3, WF_COL_LABELVAL).Value
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    AssertTrue lbl.Points(2).DataLabel.NumberFormat = "#,##0;-#,##0;" Or lbl.Points(2).DataLabel.NumberFormat = LocalizeFormat("#,##0;-#,##0;"), "formato com sinal de menos, zero vazio"
    AssertEqual xlLabelPositionBelow, lbl.Points(2).DataLabel.Position
    CloseTemp wb
End Sub

Private Sub AssertRollback(isWaterfall As Boolean)
    Dim wb As Workbook, d As clsChartData, n As Long, co As ChartObject, nm As Name, ws As Worksheet
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set ws = d.Source.Worksheet
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    If isWaterfall Then Set co = Waterfall_Insert(ws, d) Else Set co = Stacked_Insert(ws, d, "stacked")
    n = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, n, "erro original preservado"
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    If isWaterfall Then Set co = Waterfall_Insert(ws, d) Else Set co = Stacked_Insert(ws, d, "stacked")
    AssertEqual 1, ws.ChartObjects.Count
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_Insert_RollsBackOnFailure()
    AssertRollback True
End Sub

Public Sub Test_Stacked_Insert_RollsBackOnFailure()
    AssertRollback False
End Sub
