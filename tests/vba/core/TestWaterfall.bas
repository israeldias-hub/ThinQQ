Attribute VB_Name = "TestWaterfall"
Option Explicit

Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    ws.Name = "Ponte"
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
