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
    AssertEqual RGB(0, 176, 80), LINKFRAME_GREEN, "constante verde"
    AssertEqual RGB(192, 0, 0), LINKFRAME_RED, "constante vermelha"
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
' usuario e avisado para inserir ou colar de novo (revisao final da Parte D, I1). Sem moldura, nenhum aviso. 1.13.0 (X2): continua assim, a
' sonda P-U6 (TestPptSync.Test_Rebuild_SlideKeepsOldCells_PU6_WhenAvailable) mediu que manter o id nao faria o slide acompanhar.
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

' 1.13.0 (X2): cor da moldura pelo estado; sem PowerPoint (build Test: so a instancia injetada) nenhuma apresentacao usa -> vermelha.
Public Sub Test_LinkFrame_SetState_Colors()
    Dim wb As Workbook, ws As Worksheet, c As Chart, c2 As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Ano", "V"): ws.Range("A2:B2").Value = Array("T1", 1): ws.Range("A3:B3").Value = Array("T2", 2)
    ws.Range("A1:B3").Select
    Core_Insert "clustered"
    Set c = Core_ActiveChart()
    LinkFrame_Ensure c, ws.Range("A1:B3")
    AssertEqual LINKFRAME_GREEN, LinkFrame_Find(c).Line.ForeColor.RGB, "nasce verde"
    LinkFrame_SetState c, False
    AssertEqual LINKFRAME_RED, LinkFrame_Find(c).Line.ForeColor.RGB
    LinkFrame_SetState c, True
    AssertEqual LINKFRAME_GREEN, LinkFrame_Find(c).Line.ForeColor.RGB
    ws.Range("A1:B3").Select
    Core_Insert "clustered"
    Set c2 = Core_ActiveChart()
    LinkFrame_SetState c2, False                          ' grafico sem moldura: nada acontece, sem erro
    AssertTrue LinkFrame_Find(c2) Is Nothing
    Ppt_ResetTestHooks
    AssertEqual 0, Core_RefreshFrames(wb), "sem PowerPoint: nenhuma apresentacao usa -> vermelha"
    AssertEqual LINKFRAME_RED, LinkFrame_Find(c).Line.ForeColor.RGB
    AssertEqual 1, CountFrames(wb), "Core_RefreshFrames nao cria nem apaga molduras"
    AssertEqual 0, Core_RefreshFrames(Nothing), "pasta Nothing: nada"
    Teardown wb
End Sub

' Revisao final 1.13.0 (I1): PowerPoint OCUPADO ao pegar o aplicativo (gancho getapp com o erro de ocupado) = nada e pintado, a moldura
' verde continua verde e devolve 0; outro erro do GetObject = "sem PowerPoint" (vermelha), como antes.
Public Sub Test_LinkFrame_RefreshFrames_BusyKeepsColors()
    Dim wb As Workbook, ws As Worksheet, c As Chart, en As Long, ed As String
    On Error GoTo Cleanup
    Ppt_ResetTestHooks
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Ano", "V"): ws.Range("A2:B2").Value = Array("T1", 1): ws.Range("A3:B3").Value = Array("T2", 2)
    ws.Range("A1:B3").Select
    Core_Insert "clustered"
    Set c = Core_ActiveChart()
    LinkFrame_Ensure c, ws.Range("A1:B3")
    LinkFrame_SetState c, True
    PptSync_SetFailNumberForTest "getapp", -2147418111
    AssertEqual 0, Core_RefreshFrames(wb), "ocupado: nada pintado"
    AssertEqual LINKFRAME_GREEN, LinkFrame_Find(c).Line.ForeColor.RGB, "ocupado (RPC_E_CALL_REJECTED): continua verde"
    PptSync_SetFailNumberForTest "getapp", -2147417846
    AssertEqual 0, Core_RefreshFrames(wb)
    AssertEqual LINKFRAME_GREEN, LinkFrame_Find(c).Line.ForeColor.RGB, "ocupado (RPC_E_RETRYLATER): continua verde"
    PptSync_SetFailForTest "getapp"
    AssertEqual 0, Core_RefreshFrames(wb)
    AssertEqual LINKFRAME_RED, LinkFrame_Find(c).Line.ForeColor.RGB, "outro erro = sem PowerPoint: vermelha"
Cleanup:
    en = Err.Number: ed = Err.Description
    PptSync_ResetTestHooks
    If Not wb Is Nothing Then Teardown wb
    If en <> 0 Then Err.Raise en, "TestLinkFrame", ed
End Sub

' Revisao final 1.13.0 (M4): uma planilha copiada leva outra moldura com o mesmo nome (tqqlink_<id>); Core_RefreshFrames pinta CADA forma.
Public Sub Test_LinkFrame_RefreshFrames_PaintsEachShape()
    Dim wb As Workbook, rng As Range, c As Chart, ws2 As Worksheet, n As Long, s As Shape, ws As Worksheet
    Ppt_ResetTestHooks
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    LinkFrame_Ensure c, rng
    rng.Worksheet.Copy After:=rng.Worksheet
    Set ws2 = wb.Worksheets(2)
    AssertEqual 2, CountFrames(wb), "pre-condicao: a copia levou a moldura"
    For Each ws In wb.Worksheets
        For Each s In ws.Shapes
            If Left$(s.Name, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then s.Line.ForeColor.RGB = LINKFRAME_GREEN
        Next
    Next
    AssertEqual 0, Core_RefreshFrames(wb)
    For Each ws In wb.Worksheets
        For Each s In ws.Shapes
            If Left$(s.Name, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then
                AssertEqual LINKFRAME_RED, s.Line.ForeColor.RGB, "moldura em " & ws.Name
                n = n + 1
            End If
        Next
    Next
    AssertEqual 2, n
    Teardown wb
End Sub

' Revisao final 1.13.0 (M1): Mostrar/ocultar molduras sem moldura cita os nomes atuais (constante lida pelo Test-Docs).
Public Sub Test_LinkFrame_Toggle_NoFrames_Message()
    Dim wb As Workbook
    Ppt_ResetTestHooks
    Set wb = NewTempWorkbook()
    wb.Activate
    Core_ClearLastNotice
    Core_ToggleLinkFrames
    AssertEqual LINKFRAME_MSG_NONE, Core_LastNotice()
    AssertTrue InStr(LINKFRAME_MSG_NONE, "Vincular ao PowerPoint") > 0 And InStr(LINKFRAME_MSG_NONE, "Copiar como gráfico vinculado") > 0
    AssertEqual 0, InStr(LINKFRAME_MSG_NONE, "Inserir também no slide"), "sem o nome antigo"
    AssertTrue InStr(LINKFRAME_NOTE_REBUILT, "(Mais > Copiar como gráfico vinculado)") > 0, "I2: o caminho que cola o grafico recriado"
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
