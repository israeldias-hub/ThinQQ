Attribute VB_Name = "TestAxisBreak"
Option Explicit

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

Private Function Fixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Cat", "S1", "S2")
    ws.Range("A2:C2").Value = Array("a", 5, 5): ws.Range("A3:C3").Value = Array("b", 10, 10): ws.Range("A4:C4").Value = Array("c", 200, 30)
    Set Fixture = ws.Range("A1:C4")
End Function

' Onda A (item A): nas empilhadas o corte recusa um limite que deixaria o segmento de cima de uma coluna cortada inteiro acima do eixo. Um
' segmento inteiro acima do corte continua possivel quando os DADOS sobem depois do corte (o refresh oculta o rotulo dele e poe o de baixo
' na base): os testes da Tarefa 4 que precisam desse estado cortam com a coluna c em (20, 120) (piso 50) e voltam a (200, 30).
Private Sub CutThenRaise(cht As Chart, lim As Double, Optional viaCore As Boolean = True)
    LowerC cht
    If viaCore Then Core_SetAxisBreakAt lim Else Break_Set cht, lim
    RaiseC cht
End Sub

Private Sub LowerC(cht As Chart)
    cht.Parent.Parent.Range("B4").Value = 20: cht.Parent.Parent.Range("C4").Value = 120
End Sub

Private Sub RaiseC(cht As Chart)
    cht.Parent.Parent.Range("B4").Value = 200: cht.Parent.Parent.Range("C4").Value = 30
    Core_SafeRefresh cht
End Sub

Private Function YOf(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue, xlPrimary)
    YOf = cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight * (ax.MaximumScale - v) / (ax.MaximumScale - ax.MinimumScale)
End Function

Private Function XOf(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue, xlPrimary)
    XOf = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth * (v - ax.MinimumScale) / (ax.MaximumScale - ax.MinimumScale)
End Function

Private Function HasShape(cht As Chart, nm As String) As Boolean
    Dim shp As Shape
    For Each shp In cht.Shapes
        If shp.Name = nm Then HasShape = True: Exit Function
    Next
End Function

Public Sub Test_Break_Set_ClipsAxis_TotalsAtLimit_LivePos()
    Dim wb As Workbook, rng As Range, cht As Chart, blk As Range, body As Range, pc As Long, tot As Series, ax As Axis
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Set ax = cht.Axes(xlValue, xlPrimary)
    AssertTrue ax.MaximumScale >= 230, "antes: escala automatica cobre 230"
    CutThenRaise cht, 50, False
    AssertEqual "50", Registry_GetOption(cht, "vbreak")
    AssertFalse ax.MaximumScaleIsAuto: AssertEqual 50, ax.MaximumScale
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    pc = Block_FindCol(blk, "LblPos")
    AssertTrue pc > 0, "coluna LblPos"
    AssertEqual 50, CDbl(blk.Cells(1, pc).Value), "meta da coluna LblPos = BreakAt (dentro do bloco)"
    AssertEqual 10, body.Cells(2, pc).Value: AssertEqual 50, body.Cells(4, pc).Value, "LblPos = MIN(TotalPos, BreakAt)"
    AssertEqual "Pos", body.Cells(1, Stacked_PosCol(blk)).Value
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertTrue InStr(tot.Formula, body.Cells(2, pc).Resize(3, 1).Address(True, True)) > 0, "tqq_total aponta para LblPos"
    AssertEqual "230", tot.Points(3).DataLabel.Text, "o rotulo mostra o total real"
    ForceLayout cht
    AssertAlmost cht.PlotArea.InsideTop, tot.Points(3).Top, 2, "ponto do total da coluna cortada no topo do plot"
    AssertTrue tot.Points(3).DataLabel.Top >= 0 And tot.Points(3).DataLabel.Top + tot.Points(3).DataLabel.Height <= cht.PlotArea.InsideTop + 2, "rotulo acima do plot, dentro do grafico"
    AssertTrue cht.PlotArea.InsideTop >= 23.5, "folga de cima para o rotulo"
    Core_SafeRefresh cht
    AssertEqual 50, ax.MaximumScale, "refresh mantem o corte"
    Teardown wb
End Sub

' Review Focus 4: segmento atravessado, segmento acima do corte, marca de corte, e a volta quando os dados descem.
Public Sub Test_Break_SegmentCrossingLimit_LabelsAndMark()
    Dim wb As Workbook, rng As Range, cht As Chart, ws As Worksheet, s1 As Series, s2 As Series, m As Shape, p As Point
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    CutThenRaise cht, 50
    Set s1 = cht.SeriesCollection(1): Set s2 = cht.SeriesCollection(2)
    AssertEqual xlLabelPositionInsideBase, s1.Points(3).DataLabel.Position, "segmento atravessado: rotulo na base"
    AssertFalse s2.Points(3).HasDataLabel, "segmento inteiro acima do corte: sem rotulo"
    AssertEqual xlLabelPositionCenter, s1.Points(1).DataLabel.Position, "segmento abaixo do corte: centro"
    AssertTrue HasShape(cht, "tqq_ann_break_3") And HasShape(cht, "tqq_ann_break_3_w") And HasShape(cht, "tqq_ann_break_3_b"), "marca na coluna 3"
    AssertFalse HasShape(cht, "tqq_ann_break_1") Or HasShape(cht, "tqq_ann_break_2"), "sem marca nas colunas inteiras"
    ForceLayout cht
    Set m = cht.Shapes("tqq_ann_break_3"): Set p = s1.Points(3)
    AssertAlmost YOf(cht, 50) + 8, m.Top + m.Height / 2, 2.5, "marca 8 pt abaixo do topo do plot"
    AssertTrue m.Left <= p.Left + 0.5 And m.Left + m.Width >= p.Left + p.Width - 0.5, "marca atravessa a coluna"
    AssertEqual "kind=break;cat=3;part=a", m.AlternativeText
    ws.Range("B4").Value = 20: ws.Range("C4").Value = 20
    Core_SafeRefresh cht
    AssertFalse HasShape(cht, "tqq_ann_break_3"), "total 40 < 50: marca removida"
    AssertTrue s2.Points(3).HasDataLabel, "rotulo de volta"
    AssertEqual xlLabelPositionCenter, s1.Points(3).DataLabel.Position
    AssertEqual "40", Chart_FindSeries(cht, "tqq_total").Points(3).DataLabel.Text
    AssertEqual 50, cht.Axes(xlValue).MaximumScale, "o corte continua gravado"
    Teardown wb
End Sub

Public Sub Test_Break_Remove_RestoresAuto()
    Dim wb As Workbook, rng As Range, cht As Chart, blk As Range, body As Range, pc As Long
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    CutThenRaise cht, 50
    Core_RemoveAxisBreak
    AssertEqual "", Registry_GetOption(cht, "vbreak", "")
    AssertTrue cht.Axes(xlValue).MaximumScaleIsAuto
    AssertFalse HasShape(cht, "tqq_ann_break_3")
    Set blk = Registry_Block(cht): Set body = Block_Body(blk): pc = Block_FindCol(blk, "LblPos")
    AssertEqual "", CStr(blk.Cells(1, pc).Value), "limite apagado da meta da coluna LblPos"
    AssertEqual 230, body.Cells(4, pc).Value, "MIN com a meta vazia = TotalPos"
    AssertEqual xlLabelPositionCenter, cht.SeriesCollection(1).Points(3).DataLabel.Position
    AssertTrue cht.SeriesCollection(2).Points(3).HasDataLabel
    Teardown wb
End Sub

Public Sub Test_Break_Clustered_And_Bars()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, s As Series, m As Shape, tot As Series, blk As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "clustered"
    Set cht = Core_ActiveChart()
    Break_Set cht, 50
    Set s = cht.SeriesCollection(1)
    AssertEqual xlLabelPositionInsideBase, s.Points(3).DataLabel.Position, "agrupadas: ponto cortado com rotulo na base"
    AssertEqual xlLabelPositionOutsideEnd, s.Points(1).DataLabel.Position
    AssertTrue HasShape(cht, "tqq_ann_break_3")
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    rng.Select
    Core_Insert "barstacked"
    Set cht = Core_ActiveChart()
    CutThenRaise cht, 50, False
    AssertEqual 50, cht.Axes(xlValue, xlPrimary).MaximumScale
    Set blk = Registry_Block(cht)
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertTrue InStr(tot.Formula, Block_Body(blk).Cells(2, Block_FindCol(blk, "LblPos")).Resize(3, 1).Address(True, True)) > 0, "barras: X do total = LblPos"
    ForceLayout cht
    Set m = cht.Shapes("tqq_ann_break_3")
    AssertTrue m.Height > m.Width, "barras: marca atravessa a barra na vertical"
    AssertTrue cht.ChartArea.Width - (cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth) >= 20, "folga a direita para o rotulo do total"
    Teardown wb
End Sub

Public Sub Test_Break_Refusals_SameScale_Rebuild_RemoveAll()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, c2 As Chart, coll As New Collection, skipped As Long, en As Long
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "stacked100"
    Core_ClearLastNotice
    Core_SetAxisBreakAt 0.5
    AssertEqual "'Corte de eixo' não se aplica a este tipo de gráfico.", Core_LastNotice()
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_ClearLastNotice
    Core_SetAxisBreakAt 500
    AssertEqual "Corte de eixo: informe um valor maior que zero e menor que o maior valor do gráfico.", Core_LastNotice()
    AssertEqual "", Registry_GetOption(cht, "vbreak", "")
    On Error Resume Next
    Break_Set cht, -1
    en = Err.Number: Err.Clear
    On Error GoTo 0
    AssertEqual 0, en, "limite <= 0 e remover (sem erro)"
    ws.Range("A10:C12").Value = Array("x", 1, 1): ws.Range("A11:C11").Value = Array("y", 2, 2): ws.Range("A12:C12").Value = Array("z", 3, 3)
    ws.Range("A10:C12").Select
    Core_Insert "stacked"
    Set c2 = Core_ActiveChart()
    coll.Add cht: coll.Add c2
    AssertEqual 2, Deco_SameScale(coll, skipped), "mesma escala antes do corte"
    AssertTrue Len(Registry_GetOption(cht, "vmax", "")) > 0
    CutThenRaise cht, 50, False
    AssertEqual "", Registry_GetOption(cht, "vmax", ""), "o corte limpa a escala fixa"
    Set coll = New Collection: coll.Add cht: coll.Add c2
    AssertEqual 0, Deco_SameScale(coll, skipped): AssertEqual 1, skipped, "grafico com corte fica fora da mesma escala"
    AssertEqual 50, cht.Axes(xlValue).MaximumScale, "o corte nao foi desfeito"
    Ann_AddCagr cht, 1, 3
    Ann_RemoveAll cht
    AssertTrue HasShape(cht, "tqq_ann_break_3"), "Remover setas poupa a marca de corte"
    AssertFalse HasShape(cht, "tqq_ann_cagr_1")
    ws.Range("B4").Value = 20: ws.Range("C4").Value = 120         ' o piso das empilhadas (50) admite o corte nos dados do rebuild
    AssertTrue Core_RebuildFromRange(cht, rng), "rebuild"
    Set cht = Core_ActiveChart()
    AssertEqual "50", Registry_GetOption(cht, "vbreak"), "corte mantido no rebuild"
    AssertEqual 50, cht.Axes(xlValue).MaximumScale
    AssertTrue HasShape(cht, "tqq_ann_break_3")
    AssertTrue Block_FindCol(Registry_Block(cht), "LblPos") > 0
    Teardown wb
End Sub

Public Sub Test_Break_ThenRoundSum_CoexistInBlock()
    Dim wb As Workbook, rng As Range, cht As Chart, blk As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    CutThenRaise cht, 50
    Core_ClearLastNotice
    Core_SetToggle "roundsum", True
    AssertTrue Core_GetToggle("roundsum"), "soma consistente depois do corte (" & Core_LastNotice() & ")"
    Set blk = Registry_Block(cht)
    AssertTrue Round_DCol(blk) > 0 And Block_FindCol(blk, "LblPos") > 0
    AssertEqual 50, CDbl(blk.Cells(1, Block_FindCol(blk, "LblPos")).Value), "limite dentro do bloco"
    AssertEqual 50, cht.Axes(xlValue).MaximumScale
    Teardown wb
End Sub

Public Sub Test_Core_SetAxisBreak_TestBuild_NoInputBox()
    Dim wb As Workbook, rng As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked100"
    Core_ClearLastNotice
    Core_SetAxisBreak                                           ' a recusa pela matriz vem ANTES da saida do build Test
    AssertEqual "'Corte de eixo' não se aplica a este tipo de gráfico.", Core_LastNotice()
    rng.Worksheet.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    rng.Select
    Core_Insert "stacked"
    Core_ClearLastNotice
    Core_SetAxisBreak                                           ' build Test: sai antes do InputBox
    AssertEqual "", Core_LastNotice()
    AssertEqual "", Registry_GetOption(Core_ActiveChart(), "vbreak", ""), "nada aplicado sem o InputBox"
    AssertTrue Break_Suggest(Core_ActiveChart()) > 20 And Break_Suggest(Core_ActiveChart()) < 230, "sugestao entre o 2o e o maior topo"
    rng.Worksheet.Range("A1").Select
    Core_SetAxisBreak
    AssertEqual "", Core_LastNotice(), "sem grafico: nada"
    Teardown wb
End Sub

' Rotulos ocultos pelo corte ficam na lista brklbl e so eles voltam; um rotulo apagado pelo usuario continua apagado; religar Rotulos refaz
' todos e o corte oculta de novo.
Public Sub Test_Break_UserDeletedLabel_StaysDeleted_HiddenComeBack()
    Dim wb As Workbook, rng As Range, ws As Worksheet, cht As Chart, s1 As Series, s2 As Series
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    CutThenRaise cht, 50
    Set s1 = cht.SeriesCollection(1): Set s2 = cht.SeriesCollection(2)
    AssertEqual "2.3", Registry_GetOption(cht, "brklbl"), "o corte registra o rotulo que ocultou (serie.ponto)"
    s1.Points(2).HasDataLabel = False                           ' apagado pelo usuario
    Core_SafeRefresh cht
    AssertFalse s1.Points(2).HasDataLabel, "apagado pelo usuario: o refresh nao o recria"
    ws.Range("C4").Value = 31
    Core_SafeRefresh cht
    AssertFalse s2.Points(3).HasDataLabel, "continua acima do corte: oculto"
    ws.Range("B4").Value = 10: ws.Range("C4").Value = 10           ' total 20 < 50
    Core_SafeRefresh cht
    AssertTrue s2.Points(3).HasDataLabel, "o oculto pelo corte volta"
    AssertEqual "10", s2.Points(3).DataLabel.Text
    AssertEqual xlLabelPositionCenter, s2.Points(3).DataLabel.Position
    AssertEqual Style_ContrastFont(CLng(s2.Format.Fill.ForeColor.RGB)), s2.Points(3).DataLabel.Font.Color, "fonte por contraste"
    AssertEqual xlLabelPositionCenter, s1.Points(3).DataLabel.Position, "o da base volta ao centro"
    AssertFalse s1.Points(2).HasDataLabel, "o apagado pelo usuario nao volta"
    AssertEqual "", Registry_GetOption(cht, "brklbl", ""), "lista vazia"
    ws.Range("B4").Value = 200: ws.Range("C4").Value = 30
    Core_SafeRefresh cht
    Core_SetToggle "labels", False
    AssertEqual "", Registry_GetOption(cht, "brklbl", ""), "rotulos desligados: lista vazia"
    Core_ClearLastNotice
    Core_SetToggle "labels", True
    AssertEqual "", Core_LastNotice(), "religar Rotulos com o corte"
    AssertTrue s1.Points(2).HasDataLabel, "religar Rotulos refaz todos"
    AssertFalse s2.Points(3).HasDataLabel, "e o corte oculta de novo"
    AssertEqual xlLabelPositionInsideBase, s1.Points(3).DataLabel.Position
    AssertEqual "2.3", Registry_GetOption(cht, "brklbl")
    Teardown wb
End Sub

' Soma consistente ligada ANTES do corte: um rotulo ligado (Rnd_j) de ponto inteiro acima do corte recusa a cor da fonte (1004, medido); ligar e
' religar os rotulos funciona, o ponto cortado fica oculto e volta LIGADO a Rnd_j quando os dados descem.
Public Sub Test_Break_RoundSumFirst_LinkedLabelsHiddenAndBackLinked()
    Dim wb As Workbook, rng As Range, ws As Worksheet, cht As Chart, s1 As Series, s2 As Series, blk As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetToggle "roundsum", True
    Core_ClearLastNotice
    CutThenRaise cht, 50
    AssertEqual "", Core_LastNotice()
    Set s1 = cht.SeriesCollection(1): Set s2 = cht.SeriesCollection(2)
    AssertFalse s2.Points(3).HasDataLabel, "ponto cortado: oculto"
    AssertEqual xlLabelPositionInsideBase, s1.Points(3).DataLabel.Position
    ws.Range("B3").Value = 60                                   ' b: 60 + 10 (S1 atravessa, S2 inteiro acima)
    Core_SafeRefresh cht
    AssertFalse s2.Points(2).HasDataLabel, "rotulo ligado de ponto que passou do corte: oculto"
    AssertEqual xlLabelPositionInsideBase, s1.Points(2).DataLabel.Position
    ws.Range("B3").Value = 10
    Core_SafeRefresh cht
    AssertTrue s2.Points(2).HasDataLabel, "volta"
    Set blk = Registry_Block(cht)
    AssertTrue InStr(s2.Points(2).DataLabel.Formula, Block_Body(blk).Cells(3, Round_RCol(blk, 2)).Address(True, True)) > 0, "volta ligado a Rnd2 (" & s2.Points(2).DataLabel.Formula & ")"
    AssertEqual "10", s2.Points(2).DataLabel.Text
    Core_ClearLastNotice
    Core_SetToggle "labels", False
    Core_SetToggle "labels", True
    Core_SetToggle "roundsum", False
    Core_SetToggle "roundsum", True
    AssertEqual "", Core_LastNotice(), "religar com o corte e a soma consistente"
    AssertFalse s2.Points(3).HasDataLabel
    AssertTrue InStr(s1.Points(1).DataLabel.Formula, Block_Body(blk).Cells(2, Round_RCol(blk, 1)).Address(True, True)) > 0, "ligado a Rnd1"
    Teardown wb
End Sub

' Break_Set transacional: falha depois de gravar o limite (ou o eixo) volta celula, opcoes (vbreak, vmin, vmax), eixo, rotulos e marcas.
Public Sub Test_Break_Set_FailMidway_RollsBack()
    Dim wb As Workbook, rng As Range, cht As Chart, blk As Range, pc As Long, en As Long, es As String, ed As String, ax As Axis
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Set ax = cht.Axes(xlValue, xlPrimary)
    Registry_SetOption cht, "vmin", "0": Registry_SetOption cht, "vmax", "300"
    Deco_ApplyFixedScale cht
    AssertEqual 300, ax.MaximumScale
    LowerC cht
    Util_FailThinSyncForTest = "break-set:axis"
    On Error Resume Next
    Break_Set cht, 50
    en = Err.Number: Err.Clear
    RaiseC cht
    On Error GoTo Cleanup
    AssertEqual vbObjectError + 9994, en, "o erro chega ao chamador"
    AssertEqual "", Util_FailThinSyncForTest, "gancho de uso unico"
    AssertEqual "", Registry_GetOption(cht, "vbreak", "")
    AssertEqual "300", Registry_GetOption(cht, "vmax", ""), "a escala fixa volta"
    AssertEqual "0", Registry_GetOption(cht, "vmin", "")
    AssertFalse ax.MaximumScaleIsAuto: AssertEqual 300, ax.MaximumScale
    Set blk = Registry_Block(cht): pc = Block_FindCol(blk, "LblPos")
    AssertEqual "", CStr(blk.Cells(1, pc).Value), "limite apagado"
    AssertEqual 230, Block_Body(blk).Cells(4, pc).Value
    AssertFalse HasShape(cht, "tqq_ann_break_3")
    AssertTrue cht.SeriesCollection(2).Points(3).HasDataLabel
    AssertEqual "230", Chart_FindSeries(cht, "tqq_total").Points(3).DataLabel.Text, "rotulo do total"
    CutThenRaise cht, 50, False
    AssertEqual "", Registry_GetOption(cht, "vmax", "")
    LowerC cht
    Util_FailThinSyncForTest = "break-set:cell"
    On Error Resume Next
    Break_Set cht, 80
    en = Err.Number: Err.Clear
    RaiseC cht
    On Error GoTo Cleanup
    AssertEqual vbObjectError + 9994, en
    AssertEqual "50", Registry_GetOption(cht, "vbreak"), "o corte anterior fica"
    AssertEqual 50, ax.MaximumScale
    AssertEqual 50, CDbl(Registry_Block(cht).Cells(1, pc).Value)
    AssertTrue HasShape(cht, "tqq_ann_break_3")
    AssertFalse cht.SeriesCollection(2).Points(3).HasDataLabel
    ' falha logo depois de trocar o ponteiro do total (rotulos do total ja retirados): o total volta com rotulos
    Break_Remove cht
    Chart_FindSeries(cht, "tqq_total").Delete
    Stacked_SetTotals cht, True                                 ' a auxiliar nasce em LblPos; para testar a troca, volta a TotalPos
    Chart_FindSeries(cht, "tqq_total").HasDataLabels = False
    Chart_FindSeries(cht, "tqq_total").Values = Block_Ref(Block_Body(Registry_Block(cht)).Cells(2, 3).Resize(3, 1))
    Stacked_SetTotals cht, True
    LowerC cht
    Util_FailThinSyncForTest = "break-repoint:values"
    On Error Resume Next
    Break_Set cht, 50
    en = Err.Number: Err.Clear
    RaiseC cht
    On Error GoTo Cleanup
    AssertEqual vbObjectError + 9994, en
    AssertEqual "", Registry_GetOption(cht, "vbreak", "")
    AssertTrue Chart_FindSeries(cht, "tqq_total").HasDataLabels, "rotulos do total religados no caminho de erro"
    AssertEqual "230", Chart_FindSeries(cht, "tqq_total").Points(3).DataLabel.Text
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailThinSyncForTest = ""
    Teardown wb
    If en <> 0 Then Err.Raise en, es, ed
End Sub

' Break_EnsureColumn transacional: falha depois de estender limpa a coluna e devolve o nome; o corte depois funciona.
Public Sub Test_Break_EnsureColumn_FailMidway_RollsBack()
    Dim wb As Workbook, rng As Range, cht As Chart, oldRef As String, cols As Long, en As Long, es As String, ed As String, blk As Range
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    oldRef = wb.Names(ID_PREFIX & Registry_IdFromChart(cht)).RefersTo
    Set blk = Registry_Block(cht): cols = blk.Columns.Count
    LowerC cht
    Util_FailThinSyncForTest = "break-ensure:write"
    On Error Resume Next
    Break_Set cht, 50
    en = Err.Number: Err.Clear
    RaiseC cht
    On Error GoTo Cleanup
    AssertEqual vbObjectError + 9994, en
    AssertEqual oldRef, wb.Names(ID_PREFIX & Registry_IdFromChart(cht)).RefersTo, "nome devolvido"
    AssertEqual 0, Block_FindCol(Registry_Block(cht), "LblPos")
    AssertEqual 0, Application.WorksheetFunction.CountA(blk.Offset(0, cols).Resize(blk.Rows.Count, 1)), "coluna limpa"
    AssertEqual "", Registry_GetOption(cht, "vbreak", "")
    AssertTrue cht.Axes(xlValue).MaximumScaleIsAuto
    CutThenRaise cht, 50, False
    AssertEqual cols + 1, Block_FindCol(Registry_Block(cht), "LblPos"), "criada no mesmo lugar"
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailThinSyncForTest = ""
    Teardown wb
    If en <> 0 Then Err.Raise en, es, ed
End Sub

' Religar um rotulo oculto que falha no meio da formatacao: volta a ficar oculto e continua na lista; a passada seguinte o religa.
Public Sub Test_Break_Restore_FailMidway_StaysHiddenAndListed()
    Dim wb As Workbook, rng As Range, ws As Worksheet, cht As Chart, s2 As Series, en As Long, es As String, ed As String
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    CutThenRaise cht, 50
    Set s2 = cht.SeriesCollection(2)
    Application.EnableEvents = False                            ' o refresh so quando o teste chama
    ws.Range("B4").Value = 10: ws.Range("C4").Value = 10
    Util_FailThinSyncForTest = "break-restore:2.3"
    Core_SafeRefresh cht
    AssertEqual "", Util_FailThinSyncForTest, "o gancho disparou"
    AssertFalse s2.Points(3).HasDataLabel, "falha no meio: oculto de novo"
    AssertEqual "2.3", Registry_GetOption(cht, "brklbl"), "continua na lista"
    Core_SafeRefresh cht
    AssertTrue s2.Points(3).HasDataLabel, "a passada seguinte religa"
    AssertEqual "10", s2.Points(3).DataLabel.Text
    AssertEqual "", Registry_GetOption(cht, "brklbl", "")
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailThinSyncForTest = ""
    Application.EnableEvents = True
    Teardown wb
    If en <> 0 Then Err.Raise en, es, ed
End Sub

' Agrupadas: o rotulo na base fica DENTRO da coluna (fonte por contraste, refeita com a paleta) e volta cinza-escuro fora; a marca atravessa
' so as colunas cortadas da categoria (colunas e barras), pela geometria real dos pontos.
Public Sub Test_Break_Clustered_BaseContrast_MarkSpansCutBarsOnly()
    Dim wb As Workbook, rng As Range, ws As Worksheet, cht As Chart, s1 As Series, s2 As Series, m As Shape, p1 As Point, p2 As Point
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "clustered"
    Set cht = Core_ActiveChart()
    Core_SetAxisBreakAt 50
    Set s1 = cht.SeriesCollection(1): Set s2 = cht.SeriesCollection(2)
    AssertEqual Style_ContrastFont(CLng(s1.Format.Fill.ForeColor.RGB)), s1.Points(3).DataLabel.Font.Color, "base dentro da coluna: contraste"
    AssertEqual RGB(64, 64, 64), s1.Points(1).DataLabel.Font.Color
    ForceLayout cht
    Set m = cht.Shapes("tqq_ann_break_3"): Set p1 = s1.Points(3): Set p2 = s2.Points(3)
    AssertTrue m.Left <= p1.Left + 0.5 And m.Left + m.Width >= p1.Left + p1.Width - 0.5, "atravessa a coluna cortada (" & m.Left & "+" & m.Width & " / " & p1.Left & "+" & p1.Width & ")"
    AssertTrue m.Left + m.Width <= p2.Left + 3.5, "nao atravessa o espaco acima da vizinha mais baixa (" & (m.Left + m.Width) & " / " & p2.Left & ")"
    AssertAlmost YOf(cht, 50) + 8, m.Top + m.Height / 2, 2.5, "8 pt abaixo do corte"
    Core_SetOption "palette", "cinza"
    AssertEqual Style_ContrastFont(CLng(s1.Format.Fill.ForeColor.RGB)), s1.Points(3).DataLabel.Font.Color, "paleta nova: contraste refeito"
    ws.Range("B4").Value = 20
    Core_SafeRefresh cht
    AssertEqual xlLabelPositionOutsideEnd, s1.Points(3).DataLabel.Position
    AssertEqual RGB(64, 64, 64), s1.Points(3).DataLabel.Font.Color, "fora da coluna: cinza-escuro"
    AssertFalse HasShape(cht, "tqq_ann_break_3")
    ws.Range("B4").Value = 200
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    rng.Select
    Core_Insert "barclustered"
    Set cht = Core_ActiveChart()
    Core_SetAxisBreakAt 50
    Set s1 = cht.SeriesCollection(1): Set s2 = cht.SeriesCollection(2)
    ForceLayout cht
    Set m = cht.Shapes("tqq_ann_break_3"): Set p1 = s1.Points(3): Set p2 = s2.Points(3)
    AssertTrue m.Top <= p1.Top + 0.5 And m.Top + m.Height >= p1.Top + p1.Height - 0.5, "barras: atravessa a barra cortada (" & m.Top & "+" & m.Height & " / " & p1.Top & "+" & p1.Height & ")"
    AssertTrue m.Top >= p2.Top + p2.Height - 3.5 Or m.Top + m.Height <= p2.Top + 3.5, "barras: nao atravessa a vizinha mais curta (" & m.Top & "+" & m.Height & " / " & p2.Top & "+" & p2.Height & ")"
    AssertAlmost XOf(cht, 50) - 8, m.Left + m.Width / 2, 2.5, "barras: 8 pt antes do corte"
    Teardown wb
End Sub

' Barras empilhadas: marca vertical 8 pt antes do corte atravessando a barra, rotulo do total a direita do corte dentro do grafico; remover devolve.
Public Sub Test_Break_BarsStacked_MarkTotalAndRemove()
    Dim wb As Workbook, rng As Range, cht As Chart, s1 As Series, m As Shape, p As Point, tot As Series, dl As DataLabel
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "barstacked"
    Set cht = Core_ActiveChart()
    CutThenRaise cht, 50
    Set s1 = cht.SeriesCollection(1): Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertEqual xlLabelPositionInsideBase, s1.Points(3).DataLabel.Position
    AssertFalse cht.SeriesCollection(2).Points(3).HasDataLabel
    AssertEqual "230", tot.Points(3).DataLabel.Text
    ForceLayout cht
    Set m = cht.Shapes("tqq_ann_break_3"): Set p = s1.Points(3)
    AssertAlmost XOf(cht, 50) - 8, m.Left + m.Width / 2, 2.5, "marca 8 pt antes do corte"
    AssertTrue m.Top <= p.Top + 0.5 And m.Top + m.Height >= p.Top + p.Height - 0.5, "atravessa a barra"
    Set dl = tot.Points(3).DataLabel
    AssertTrue dl.Left >= cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth - 1, "total a direita do corte (" & dl.Left & ")"
    AssertTrue dl.Left + dl.Width <= Geom_ChartRightEdge(cht) + 1, "antes da legenda"
    Core_RemoveAxisBreak
    AssertTrue cht.Axes(xlValue).MaximumScaleIsAuto
    AssertFalse HasShape(cht, "tqq_ann_break_3")
    AssertTrue cht.SeriesCollection(2).Points(3).HasDataLabel
    AssertEqual xlLabelPositionCenter, s1.Points(3).DataLabel.Position
    AssertEqual "230", tot.Points(3).DataLabel.Text
    Teardown wb
End Sub

' Comando com a resposta do InputBox pelo gancho de teste (uso unico, consumido ja na entrada): valor <= 0, Cancelar e valor valido.
Public Sub Test_Core_SetAxisBreak_AnswerHook()
    Dim wb As Workbook, rng As Range, cht As Chart
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked100"
    Core_TestSetAxisBreakAnswer 50
    Core_ClearLastNotice
    Core_SetAxisBreak
    AssertEqual "'Corte de eixo' não se aplica a este tipo de gráfico.", Core_LastNotice()
    rng.Worksheet.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_ClearLastNotice
    Core_SetAxisBreak
    AssertEqual "", Registry_GetOption(cht, "vbreak", ""), "o gancho foi consumido na recusa (nao vaza)"
    Core_TestSetAxisBreakAnswer 0
    Core_SetAxisBreak
    AssertEqual "Corte de eixo: informe um valor maior que zero e menor que o maior valor do gráfico.", Core_LastNotice()
    AssertEqual "", Registry_GetOption(cht, "vbreak", "")
    Core_ClearLastNotice
    Core_TestSetAxisBreakAnswer False
    Core_SetAxisBreak
    AssertEqual "", Core_LastNotice(), "Cancelar"
    AssertEqual "", Registry_GetOption(cht, "vbreak", "")
    Core_TestSetAxisBreakAnswer 50
    Core_SetAxisBreak
    AssertTrue InStr(Core_LastNotice(), "segmento de cima") > 0, "abaixo do piso das empilhadas: aviso (" & Core_LastNotice() & ")"
    AssertEqual "", Registry_GetOption(cht, "vbreak", "")
    LowerC cht
    Core_TestSetAxisBreakAnswer 50
    Core_SetAxisBreak
    AssertEqual "50", Registry_GetOption(cht, "vbreak", "")
    AssertEqual 50, cht.Axes(xlValue).MaximumScale
    Teardown wb
End Sub

' Totais ligados depois do corte: a auxiliar nasce em LblPos e a folga de cima vem no refresh. "Alterar dados" para uma tabela cujo maior topo
' fica abaixo do corte: o corte e desfeito (registrado), sem marca e com o eixo automatico.
Public Sub Test_Break_TotalsAfterBreak_RebuildBelowLimit()
    Dim wb As Workbook, rng As Range, ws As Worksheet, cht As Chart, tot As Series, blk As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetToggle "totals", False
    CutThenRaise cht, 50
    Core_SetToggle "totals", True
    Set tot = Chart_FindSeries(cht, "tqq_total"): Set blk = Registry_Block(cht)
    AssertTrue InStr(tot.Formula, Block_Body(blk).Cells(2, Block_FindCol(blk, "LblPos")).Resize(3, 1).Address(True, True)) > 0, "nasce em LblPos"
    AssertEqual "230", tot.Points(3).DataLabel.Text
    ForceLayout cht
    AssertTrue cht.PlotArea.InsideTop >= 23.5, "folga do total ligado depois do corte (" & cht.PlotArea.InsideTop & ")"
    AssertTrue tot.Points(3).DataLabel.Top >= 0, "rotulo dentro do grafico"
    ws.Range("A10:C10").Value = Array("Cat", "S1", "S2"): ws.Range("A11:C11").Value = Array("x", 1, 2): ws.Range("A12:C12").Value = Array("y", 3, 4)
    Core_ClearLastNotice
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A10:C12")), "rebuild"
    AssertTrue InStr(Core_LastNotice(), "Corte de eixo retirado") > 0, "retirado com aviso (" & Core_LastNotice() & ")"
    Set cht = Core_ActiveChart()
    AssertEqual "", Registry_GetOption(cht, "vbreak", ""), "limite acima dos dados novos: corte desfeito"
    AssertTrue cht.Axes(xlValue).MaximumScaleIsAuto
    AssertFalse HasShape(cht, "tqq_ann_break_1") Or HasShape(cht, "tqq_ann_break_2")
    AssertTrue cht.SeriesCollection(2).Points(2).HasDataLabel
    Teardown wb
End Sub

' Mesma escala com um grafico com corte: o aviso diz por que ele fica de fora.
Public Sub Test_Break_SameScaleNotice_MentionsBreak()
    Dim wb As Workbook, rng As Range, ws As Worksheet
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "stacked"
    CutThenRaise Core_ActiveChart(), 50
    ws.Range("A10:C10").Value = Array("Cat", "S1", "S2"): ws.Range("A11:C11").Value = Array("x", 1, 2): ws.Range("A12:C12").Value = Array("y", 3, 4)
    ws.Range("A10:C12").Select
    Core_Insert "stacked"
    ws.Range("A1").Select
    Core_ClearLastNotice
    Core_SameScale
    AssertEqual "Mesma escala: são necessários dois ou mais gráficos com escala livre (100%, Mekko (%), pizza e Gantt não entram). Gráficos com corte de eixo também ficam de fora.", Core_LastNotice()
    Teardown wb
End Sub

' Regressao (medido): trocar Values da auxiliar tqq_total com os rotulos ligados passava as ligacoes para rotulos de OUTRA serie (o segmento
' Online de 2025, valor 13, mostrava "30", o total de 2022). Todo rotulo de segmento visivel mostra o proprio valor, em colunas e barras.
Public Sub Test_Break_SegmentLabelsKeepTheirValues()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, s As Series, i As Long, k As Variant, v As Variant, vis As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:D1").Value = Array("Ano", "Lojas", "Online", "Atacado")
    ws.Range("A2:D2").Value = Array("2021", 12, 8, 5): ws.Range("A3:D3").Value = Array("2022", 15, 9, 6)
    ws.Range("A4:D4").Value = Array("2023", 18, 11, 7): ws.Range("A5:D5").Value = Array("2024", 160, 40, 9)
    ws.Range("A6:D6").Value = Array("2025", 21, 13, 8)
    Set rng = ws.Range("A1:D6")
    For Each k In Array("stacked", "barstacked")
        rng.Select
        Core_Insert CStr(k)
        Set cht = Core_ActiveChart()
        ws.Range("B5:D5").Value = Array(30, 10, 9)              ' piso das empilhadas em 45: 49 - 0,75 x 9 = 42,25
        Core_SetAxisBreakAt 45
        ws.Range("B5:D5").Value = Array(160, 40, 9)             ' os dados sobem depois do corte: Online e Atacado de 2024 inteiros acima
        Core_SafeRefresh cht
        vis = 0
        For Each s In cht.SeriesCollection
            If Not Style_IsHelperSeries(s) Then
                v = s.Values
                For i = 1 To s.Points.Count
                    If s.Points(i).HasDataLabel Then
                        vis = vis + 1
                        AssertEqual CStr(v(LBound(v) + i - 1)), s.Points(i).DataLabel.Text, k & ": " & s.Name & " ponto " & i
                        AssertTrue InStr(s.Points(i).DataLabel.Formula, "_ThinQQ") = 0, k & ": " & s.Name & " ponto " & i & " sem ligacao a planilha interna"
                    End If
                Next
            End If
        Next
        AssertEqual 13, vis, k & ": 15 segmentos menos os 2 acima do corte"
        AssertEqual "209", Chart_FindSeries(cht, "tqq_total").Points(4).DataLabel.Text
        ws.ChartObjects(1).Delete
        CalcSheet_Cleanup wb
    Next
    Teardown wb
End Sub

' Onda A, item A: nas empilhadas o corte nunca deixa o segmento de cima de uma coluna cortada inteiro acima do eixo (na demonstracao a coluna
' do cimento parecia 100% B2C). Piso = base do segmento de cima + 25% dele (29.255 + 0,25 x 14.068 = 32.772): a sugestao respeita o piso e
' um limite abaixo dele e recusado com aviso (nada gravado). Agrupadas nao tem piso (cada serie e uma coluna inteira).
Public Sub Test_Break_Stacked_TopSegmentFloor_DemoTable()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, k As Variant, s2 As Series, sug As Double, en As Long, ed As String
    Dim fl As Double, vis As Double, ax As Axis
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("Departamento", "B2C", "B2B")
    ws.Range("A2:C2").Value = Array("Cimento", 29255, 14068): ws.Range("A3:C3").Value = Array("Tintas", 12000, 5000)
    ws.Range("A4:C4").Value = Array("Hidraulica", 9000, 4000): ws.Range("A5:C5").Value = Array("Eletrica", 7000, 3000)
    Set rng = ws.Range("A1:C5")
    fl = 29255 + 0.25 * 14068
    For Each k In Array("stacked", "barstacked")
        rng.Select
        Core_Insert CStr(k)
        Set cht = Core_ActiveChart()
        sug = Break_Suggest(cht)
        AssertTrue sug >= fl And sug < 43323, k & ": sugestao acima do piso do cimento (" & sug & ")"
        AssertEqual 33000, sug, k & ": piso arredondado para cima a 2 algarismos"
        Core_ClearLastNotice
        Core_SetAxisBreakAt 29000
        AssertTrue InStr(Core_LastNotice(), "Cimento") > 0 And InStr(Core_LastNotice(), "segmento de cima") > 0, k & ": aviso cita a coluna (" & Core_LastNotice() & ")"
        AssertTrue InStr(Core_LastNotice(), Format$(33000, "#,##0")) > 0, k & ": aviso traz o minimo (" & Core_LastNotice() & ")"
        AssertEqual "", Registry_GetOption(cht, "vbreak", ""), k & ": nada gravado"
        AssertTrue cht.Axes(xlValue, xlPrimary).MaximumScaleIsAuto, k & ": eixo automatico"
        On Error Resume Next
        Break_Set cht, 32000
        en = Err.Number: ed = Err.Description: Err.Clear
        On Error GoTo 0
        AssertEqual vbObjectError + 7003, en, k & ": Break_Set recusa abaixo do piso (" & ed & ")"
        Core_ClearLastNotice
        Core_SetAxisBreakAt sug
        AssertEqual "", Core_LastNotice(), k & ": a sugestao e aceita"
        Set s2 = cht.SeriesCollection(2)
        AssertTrue s2.Points(1).HasDataLabel, k & ": B2B do cimento rotulado"
        AssertEqual xlLabelPositionInsideBase, s2.Points(1).DataLabel.Position, k & ": B2B atravessado: rotulo na base"
        AssertEqual "", Registry_GetOption(cht, "brklbl", ""), k & ": nenhum rotulo oculto"
        ' regra de tres no eixo real: a base do B2B (topo do B2C) fica dentro da area de plotagem, a pelo menos 25% do segmento da borda do
        ' corte; a base desenhada do ponto confere com a escala (o Excel nao recorta Point.Height: a base e que vale, medido na Tarefa 4)
        Set ax = cht.Axes(xlValue, xlPrimary)
        ForceLayout cht
        If k = "stacked" Then
            vis = YOf(cht, 29255) - cht.PlotArea.InsideTop
            AssertAlmost YOf(cht, 29255), s2.Points(1).Top + s2.Points(1).Height, 1.5, k & ": base do B2B na escala real"
            AssertTrue vis >= 0.25 * 14068 * cht.PlotArea.InsideHeight / (ax.MaximumScale - ax.MinimumScale) - 0.5, k & ": B2B visivel (" & vis & " pt)"
        Else
            vis = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth - XOf(cht, 29255)
            AssertAlmost XOf(cht, 29255), s2.Points(1).Left, 1.5, k & ": base do B2B na escala real"
            AssertTrue vis >= 0.25 * 14068 * cht.PlotArea.InsideWidth / (ax.MaximumScale - ax.MinimumScale) - 0.5, k & ": B2B visivel (" & vis & " pt)"
        End If
        ws.ChartObjects(1).Delete
        CalcSheet_Cleanup wb
    Next
    rng.Select
    Core_Insert "clustered"
    Set cht = Core_ActiveChart()
    Core_ClearLastNotice
    Core_SetAxisBreakAt 20000
    AssertEqual "", Core_LastNotice(), "agrupadas: sem piso"
    AssertEqual "20000", Registry_GetOption(cht, "vbreak", "")
    Teardown wb
End Sub

' Onda A, rodada 1 (Importante 1): numa empilhada de UMA serie (ou categoria com base 0) o unico segmento sempre atravessa o corte: nada pode
' sumir e nao ha piso. 10 / 12 / 200 aceita o corte em 15 e a sugestao volta a 1,2 x o 2o maior topo (14,4 -> 14). A tabela da demonstracao
' continua recusada abaixo do piso, com o texto novo (segmento de cima quase invisivel) e o minimo.
Public Sub Test_Break_Stacked_SingleSeries_NoFloor()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:B1").Value = Array("Cat", "S1")
    ws.Range("A2:B2").Value = Array("a", 10): ws.Range("A3:B3").Value = Array("b", 12): ws.Range("A4:B4").Value = Array("c", 200)
    ws.Range("A1:B4").Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    AssertEqual 14, Break_Suggest(cht), "sugestao de antes (1,2 x 12 = 14,4 -> 14)"
    Core_ClearLastNotice
    Core_SetAxisBreakAt 15
    AssertEqual "", Core_LastNotice(), "uma serie: sem piso"
    AssertEqual "15", Registry_GetOption(cht, "vbreak", "")
    AssertEqual xlLabelPositionInsideBase, cht.SeriesCollection(1).Points(3).DataLabel.Position, "o segmento unico atravessa: rotulo na base"
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    ws.Range("A1:C1").Value = Array("Departamento", "B2C", "B2B")
    ws.Range("A2:C2").Value = Array("Cimento", 29255, 14068): ws.Range("A3:C3").Value = Array("Tintas", 12000, 5000)
    ws.Range("A4:C4").Value = Array("Hidraulica", 9000, 4000)
    ws.Range("A1:C4").Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_ClearLastNotice
    Core_SetAxisBreakAt 30000                                   ' acima da base (29.255) e abaixo do piso: o B2B quase nao apareceria
    AssertTrue InStr(Core_LastNotice(), "quase invisível") > 0 And InStr(Core_LastNotice(), "Cimento") > 0, "texto novo (" & Core_LastNotice() & ")"
    AssertTrue InStr(Core_LastNotice(), "inteiro acima") = 0, "sem dizer que some inteiro"
    AssertTrue InStr(Core_LastNotice(), Format$(33000, "#,##0")) > 0, "o minimo"
    AssertEqual "", Registry_GetOption(cht, "vbreak", "")
    Teardown wb
End Sub

' Onda A, rodada 1: "Alterar dados" reaplica o corte como o refresh: com os dados novos abaixo do piso (o segmento de cima da coluna c inteiro
' acima do corte), o corte fica, o segmento fica sem rotulo e entra em brklbl — o mesmo estado de quando os dados sobem depois do corte.
' Rodada 2: e um aviso (um so) diz a coluna e o minimo que mostraria o segmento.
Public Sub Test_Break_Rebuild_BelowFloor_KeptLikeRefresh()
    Dim wb As Workbook, rng As Range, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    LowerC cht
    Core_SetAxisBreakAt 50
    ws.Range("B4").Value = 200: ws.Range("C4").Value = 30        ' sem refresh: o Alterar dados ve os dados novos
    Core_ClearLastNotice
    AssertTrue Core_RebuildFromRange(cht, rng), "rebuild"
    Set cht = Core_ActiveChart()
    AssertTrue InStr(Core_LastNotice(), "'c'") > 0 And InStr(Core_LastNotice(), "segmento de cima") > 0, "aviso: o corte esconde o segmento de cima da coluna c (" & Core_LastNotice() & ")"
    AssertTrue InStr(Core_LastNotice(), Format$(210, "#,##0")) > 0, "e o minimo para ele aparecer"
    AssertEqual 1, Core_NoticeCount(), "um aviso so"
    AssertEqual "50", Registry_GetOption(cht, "vbreak", ""), "o corte fica"
    AssertFalse cht.SeriesCollection(2).Points(3).HasDataLabel, "segmento inteiro acima do corte: sem rotulo"
    AssertEqual "2.3", Registry_GetOption(cht, "brklbl", ""), "e registrado, como no refresh"
    AssertEqual xlLabelPositionInsideBase, cht.SeriesCollection(1).Points(3).DataLabel.Position
    Teardown wb
End Sub

' Onda A, rodada 2 (menor 3): o aviso do corte em "Alterar dados" sai DEPOIS de o grafico antigo ser trocado. Uma falha depois de reaplicar as
' opcoes (gancho) descarta o grafico novo e mantem o antigo com o corte: nenhum aviso de "corte retirado" sobre um grafico que continua cortado.
Public Sub Test_Break_Rebuild_NoticeOnlyAfterSwap()
    Dim wb As Workbook, rng As Range, ws As Worksheet, cht As Chart, en As Long, es As String, ed As String
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb): Set ws = rng.Worksheet
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    CutThenRaise cht, 50
    ws.Range("A10:C10").Value = Array("Cat", "S1", "S2"): ws.Range("A11:C11").Value = Array("x", 1, 2): ws.Range("A12:C12").Value = Array("y", 3, 4)
    Util_FailThinSyncForTest = "rebuild-afterreapply:1"
    Core_ClearLastNotice
    AssertFalse Core_RebuildFromRange(cht, ws.Range("A10:C12")), "falha simulada depois de reaplicar"
    AssertEqual "", Util_FailThinSyncForTest, "o gancho disparou"
    AssertTrue InStr(Core_LastNotice(), "Corte de eixo retirado") = 0, "sem aviso de corte retirado (" & Core_LastNotice() & ")"
    AssertEqual 1, Core_NoticeCount(), "so o aviso da falha (o de corte retirado nao saiu antes dela)"
    AssertEqual 1, ws.ChartObjects.Count, "o antigo fica"
    AssertEqual "50", Registry_GetOption(ws.ChartObjects(1).Chart, "vbreak", ""), "e continua cortado"
    Core_ClearLastNotice
    AssertTrue Core_RebuildFromRange(ws.ChartObjects(1).Chart, ws.Range("A10:C12")), "sem a falha"
    AssertTrue InStr(Core_LastNotice(), "Corte de eixo retirado") > 0, "agora o aviso (" & Core_LastNotice() & ")"
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailThinSyncForTest = ""
    Teardown wb
    If en <> 0 Then Err.Raise en, es, ed
End Sub
