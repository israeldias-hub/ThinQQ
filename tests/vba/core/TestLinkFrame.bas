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

' Alterar dados recria o grafico com id NOVO (o slide deixa de acompanha-lo): a moldura NAO e redesenhada, a do grafico antigo sai e o
' usuario e avisado para inserir ou colar de novo (revisao final da Parte D, I1). Sem moldura, nenhum aviso.
Public Sub Test_LinkFrame_ChangeSource_DropsFrameAndNotifies()
    Dim wb As Workbook, rng As Range, c As Chart, s As Shape, ws As Worksheet, nc As Chart
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    Set ws = rng.Worksheet
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Set s = LinkFrame_Ensure(c, rng)
    AssertTrue Not s Is Nothing
    Core_ClearLastNotice
    AssertTrue Core_RebuildFromRange(c, ws.Range("A1:B3")), "Alterar dados refez o grafico"
    Set nc = Core_ActiveChart()
    AssertTrue LinkFrame_Find(nc) Is Nothing, "o grafico novo nao ganha moldura"
    AssertEqual 0, CountFrames(wb), "a moldura do grafico antigo nao sobra"
    AssertEqual LINKFRAME_NOTE_REBUILT, Core_LastNotice(), "aviso: o slide nao acompanha mais"
    Core_ClearLastNotice
    AssertTrue Core_RebuildFromRange(nc, ws.Range("A1:C3")), "de novo, sem moldura"
    AssertTrue Core_LastNotice() <> LINKFRAME_NOTE_REBUILT, "sem moldura: sem o aviso"
    AssertEqual 0, CountFrames(wb)
    Teardown wb
End Sub

' Revisao final M2: a ativacao (Watchers_Rehook) limpa so a planilha ativada; Core_Refresh varre a pasta inteira.
Public Sub Test_LinkFrame_CleanOrphans_SheetScopedOnActivate()
    Dim wb As Workbook, rng As Range, ws1 As Worksheet, ws2 As Worksheet, c As Chart, s As Shape
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    Set ws1 = rng.Worksheet
    Set ws2 = wb.Worksheets.Add(After:=ws1)
    ws1.Activate
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Set s = LinkFrame_Ensure(c, rng)
    AssertTrue Not s Is Nothing
    ws1.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10).Name = LINKFRAME_PREFIX & "naoexiste1"
    ws2.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10).Name = LINKFRAME_PREFIX & "naoexiste2"
    AssertEqual 3, CountFrames(wb)
    LinkFrame_CleanOrphansSheet ws2
    AssertEqual 2, CountFrames(wb), "so a orfa da planilha pedida sai"
    ws2.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10).Name = LINKFRAME_PREFIX & "naoexiste3"
    ws1.Activate
    Watchers_Rehook
    AssertEqual 2, CountFrames(wb), "ativacao: a orfa da planilha ativada sai, a da outra planilha fica"
    AssertTrue Not LinkFrame_Find(c) Is Nothing, "a moldura de grafico existente fica"
    ws1.Range("A1").Select
    Core_Refresh
    AssertEqual 1, CountFrames(wb), "Core_Refresh: varre a pasta inteira"
    AssertTrue Not LinkFrame_Find(c) Is Nothing
    Teardown wb
End Sub

' Revisao final M4: falha depois do AddShape nao deixa retangulo (azul, sem nome tqqlink_) na planilha.
Public Sub Test_LinkFrame_DrawFailure_LeavesNoShape()
    Dim wb As Workbook, rng As Range, c As Chart, n0 As Long, en As Long, ed As String
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    n0 = rng.Worksheet.Shapes.Count
    LinkFrame_SetFailDrawForTest True
    AssertTrue LinkFrame_Ensure(c, rng) Is Nothing, "falhou: Nothing"
    AssertEqual n0, rng.Worksheet.Shapes.Count, "nenhuma forma sobra"
    AssertTrue Not LinkFrame_Ensure(c, rng) Is Nothing, "o gancho foi zerado no uso"
Cleanup:
    en = Err.Number: ed = Err.Description
    LinkFrame_SetFailDrawForTest False
    If Not wb Is Nothing Then Teardown wb
    If en <> 0 Then Err.Raise en, "TestLinkFrame", ed
End Sub

Private Function CountFrames(wb As Workbook) As Long
    Dim ws As Worksheet, s As Shape
    For Each ws In wb.Worksheets
        For Each s In ws.Shapes
            If Left$(s.Name, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then CountFrames = CountFrames + 1
        Next
    Next
End Function
