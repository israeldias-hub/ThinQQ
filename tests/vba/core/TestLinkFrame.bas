Attribute VB_Name = "TestLinkFrame"
Option Explicit
' 1.7.0: moldura verde em volta do intervalo de origem de um grafico que foi ao PowerPoint (spec 2026-10-15 secao 5.2).

Private Function Fixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5): ws.Range("A3:C3").Value = Array("T2", 20, 6)
    Set Fixture = ws.Range("A1:C3")
End Function

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

Public Sub Test_LinkFrame_DrawnAroundRangeAndFollowsRows()
    Dim wb As Workbook, rng As Range, c As Chart, s As Shape
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Set s = LinkFrame_Ensure(c, rng)
    AssertTrue Not s Is Nothing
    AssertEqual LINKFRAME_PREFIX & Registry_IdFromChart(c), s.Name
    AssertAlmost rng.Top, s.Top, 0.5: AssertAlmost rng.Left, s.Left, 0.5
    AssertAlmost rng.Height, s.Height, 0.5: AssertAlmost rng.Width, s.Width, 0.5
    AssertEqual msoFalse, s.Fill.Visible, "sem preenchimento"
    AssertEqual RGB(0, 176, 80), s.Line.ForeColor.RGB
    AssertEqual xlMoveAndSize, s.Placement
    rng.Worksheet.Rows(1).Insert
    AssertAlmost rng.Worksheet.Range("A2:C4").Top, s.Top, 0.5, "linha acima: a moldura desce"
    rng.Worksheet.Rows(3).Insert
    AssertAlmost rng.Worksheet.Range("A2:C5").Height, s.Height, 0.5, "linha dentro: a moldura cresce"
    Set s = LinkFrame_Ensure(c, rng.Worksheet.Range("A2:B5"))
    AssertEqual 1, CountFrames(wb), "refazer nao duplica"
    LinkFrame_Remove c
    AssertEqual 0, CountFrames(wb)
    Teardown wb
End Sub

Public Sub Test_LinkFrame_OrphansAndToggle()
    Dim wb As Workbook, rng As Range, c As Chart, s As Shape
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Set s = LinkFrame_Ensure(c, rng)
    AssertEqual 0, LinkFrame_ToggleAll(wb), "visivel -> oculta"
    AssertEqual msoFalse, s.Visible
    AssertEqual 1, LinkFrame_ToggleAll(wb), "oculta -> visivel"
    c.Parent.Delete
    LinkFrame_CleanOrphans wb
    AssertEqual 0, CountFrames(wb), "grafico apagado: moldura orfa apagada"
    AssertEqual -1, LinkFrame_ToggleAll(wb), "sem molduras"
    Teardown wb
End Sub

Public Sub Test_LinkFrame_FollowsChangeSource()
    Dim wb As Workbook, rng As Range, c As Chart, s As Shape, ws As Worksheet, nc As Chart
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    Set ws = rng.Worksheet
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Set s = LinkFrame_Ensure(c, rng)
    AssertTrue Not s Is Nothing
    AssertTrue Core_RebuildFromRange(c, ws.Range("A1:B3")), "Alterar dados refez o grafico"
    Set nc = Core_ActiveChart()
    Set s = LinkFrame_Find(nc)
    AssertTrue Not s Is Nothing, "o grafico novo tem moldura"
    AssertAlmost ws.Range("A1:B3").Width, s.Width, 0.5, "a moldura acompanha a tabela nova"
    AssertAlmost ws.Range("A1:B3").Height, s.Height, 0.5
    AssertEqual 1, CountFrames(wb), "a moldura do grafico antigo nao sobra"
    Teardown wb
End Sub

Private Function CountFrames(wb As Workbook) As Long
    Dim ws As Worksheet, s As Shape
    For Each ws In wb.Worksheets
        For Each s In ws.Shapes
            If Left$(s.Name, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then CountFrames = CountFrames + 1
        Next
    Next
End Function
