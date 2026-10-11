Attribute VB_Name = "TestPie"
Option Explicit
' Pizza e rosca (Fase 2, Tarefa 8): rotulos nativos (valor ou participacao), dentro/fora (so na pizza), cores por fatia.

Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:B1").Value = Array("Produto", "Vendas")
    ws.Range("A2:B2").Value = Array("Eletro", 50)
    ws.Range("A3:B3").Value = Array("Móveis", 30)
    ws.Range("A4:B4").Value = Array("Tintas", 20)
    Set Fixture = ChartData_FromRange(ws.Range("A1:B4"))
End Function

Private Function Dist(cht As Chart, dl As DataLabel) As Double
    Dim cx As Double, cy As Double
    cx = cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth / 2
    cy = cht.PlotArea.InsideTop + cht.PlotArea.InsideHeight / 2
    Dist = Sqr((dl.Left + dl.Width / 2 - cx) ^ 2 + (dl.Top + dl.Height / 2 - cy) ^ 2)
End Function

Private Function Radius(cht As Chart) As Double
    Radius = IIf(cht.PlotArea.InsideWidth < cht.PlotArea.InsideHeight, cht.PlotArea.InsideWidth, cht.PlotArea.InsideHeight) / 2
End Function

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

' Conta so os nomes definidos do ThinQQ (o Excel pode manter outros nomes no livro).
Private Function TqqNameCount(wb As Workbook) As Long
    Dim nm As Name
    For Each nm In wb.Names
        If Left$(nm.Name, Len(ID_PREFIX)) = ID_PREFIX Then TqqNameCount = TqqNameCount + 1
    Next
End Function

Public Sub Test_Pie_Insert_SeriesColorsLabelsLegend()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, pal As Variant, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Pie_Insert(d.Source.Worksheet, d, "pie").Chart
    AssertEqual "pie", Registry_ChartType(cht)
    AssertEqual xlPie, cht.ChartType
    AssertEqual 1, cht.SeriesCollection.Count
    Set s = cht.SeriesCollection(1)
    AssertEqual "Vendas", s.Name
    AssertEqual 0, cht.ChartGroups(1).FirstSliceAngle
    pal = Style_Palette("padrao")
    AssertEqual CLng(pal(0)), s.Points(1).Format.Fill.ForeColor.RGB
    AssertEqual CLng(pal(2)), s.Points(3).Format.Fill.ForeColor.RGB
    AssertTrue s.HasDataLabels
    AssertEqual xlLabelPositionInsideEnd, s.DataLabels.Position
    AssertTrue s.DataLabels.ShowValue: AssertFalse s.DataLabels.ShowPercentage: AssertFalse s.DataLabels.ShowCategoryName
    AssertEqual "50", s.Points(1).DataLabel.Text
    AssertTrue cht.HasLegend: AssertEqual 3, cht.Legend.LegendEntries.Count
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual 50, body.Cells(2, 2).Value: AssertAlmost 0.5, body.Cells(2, 3).Value
    AssertEqual "0", Registry_GetOption(cht, "pieout"): AssertEqual "3", Registry_GetOption(cht, "n")
    AssertEqual "right", Registry_GetOption(cht, "legend")
    Pie_SetLabels cht, False
    AssertFalse s.HasDataLabels: AssertEqual "0", Registry_GetOption(cht, "labels")
    CloseTemp wb
End Sub

Public Sub Test_Pie_NumFmtPct_ShowsPercentage()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Pie_Insert(d.Source.Worksheet, d, "pie").Chart
    Set s = cht.SeriesCollection(1)
    Pie_ApplyNumberFormat cht, "pct0"
    AssertTrue s.DataLabels.ShowPercentage: AssertFalse s.DataLabels.ShowValue
    AssertEqual "50%", s.Points(1).DataLabel.Text: AssertEqual "20%", s.Points(3).DataLabel.Text
    Pie_ApplyNumberFormat cht, "dec1"
    AssertEqual "50" & Application.International(xlDecimalSeparator) & "0", s.Points(1).DataLabel.Text
    CloseTemp wb
End Sub

Public Sub Test_Pie_Outside_CategoryNamesNoLegend_AndGeometry()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Pie_Insert(d.Source.Worksheet, d, "pie").Chart
    Set s = cht.SeriesCollection(1)
    ForceLayout cht
    AssertTrue Dist(cht, s.Points(1).DataLabel) < Radius(cht), "rotulo interno dentro do circulo"
    Pie_SetOutside cht, True
    AssertEqual "1", Registry_GetOption(cht, "pieout")
    AssertEqual xlLabelPositionOutsideEnd, s.DataLabels.Position
    AssertTrue s.DataLabels.ShowCategoryName
    AssertTrue InStr(s.Points(1).DataLabel.Text, "Eletro") > 0 And InStr(s.Points(1).DataLabel.Text, "50") > 0, s.Points(1).DataLabel.Text
    AssertFalse cht.HasLegend
    ForceLayout cht
    AssertTrue Dist(cht, s.Points(1).DataLabel) > Radius(cht), "rotulo externo fora do circulo"
    Deco_SetLegend cht, "right"
    AssertFalse cht.HasLegend, "rotulos fora vencem a opcao de legenda"
    AssertEqual "right", Registry_GetOption(cht, "legend")
    Pie_SetOutside cht, False
    AssertTrue cht.HasLegend: AssertFalse s.DataLabels.ShowCategoryName
    AssertEqual xlLabelPositionInsideEnd, s.DataLabels.Position
    Deco_SetLegend cht, "none"
    Pie_SetOutside cht, True: Pie_SetOutside cht, False
    AssertFalse cht.HasLegend, "desligar rotulos fora respeita legend=none"
    CloseTemp wb
End Sub

Public Sub Test_Doughnut_HoleAndKind()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Pie_Insert(d.Source.Worksheet, d, "doughnut").Chart
    AssertEqual "doughnut", Registry_ChartType(cht)
    AssertEqual xlDoughnut, cht.ChartType
    AssertEqual 50, cht.ChartGroups(1).DoughnutHoleSize
    AssertEqual "30", cht.SeriesCollection(1).Points(2).DataLabel.Text
    CloseTemp wb
End Sub

' Sonda (Excel 2019): a rosca recusa DataLabels.Position InsideEnd/OutsideEnd (erro 1004). Por isso a rosca nunca
' define Position e nao tem 'pieout'. Se esta sonda falhar, o Excel mudou: revise a regra (nao "conserte" o teste).
Public Sub Test_Doughnut_RejectsLabelPosition_Probe_And_NoPieOut()
    Dim wb As Workbook, d As clsChartData, cht As Chart, en As Long, ed As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Pie_Insert(d.Source.Worksheet, d, "doughnut").Chart
    AssertTrue cht.SeriesCollection(1).HasDataLabels, "rosca com rotulos na posicao padrao"
    On Error Resume Next
    cht.SeriesCollection(1).DataLabels.Position = xlLabelPositionOutsideEnd
    en = Err.Number: Err.Clear
    On Error GoTo 0
    AssertTrue en <> 0, "sonda: o Excel aceitou OutsideEnd na rosca"
    On Error Resume Next
    Pie_SetOutside cht, True
    en = Err.Number: ed = Err.Description: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    AssertEqual "'Rótulos fora da pizza' não se aplica a este tipo de gráfico.", ed
    AssertEqual "", Registry_GetOption(cht, "pieout", ""), "rosca nao grava pieout"
    AssertTrue cht.HasLegend
    CloseTemp wb
End Sub

' Review Focus 5: negativos, zero e texto viram 0 (sem tocar na tabela), 7 fatias ciclam a paleta, rotulos externos cabem.
Public Sub Test_Pie_NegativeZeroText_AndSevenSlices()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, s As Series, body As Range, i As Long, pal As Variant, dl As DataLabel
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:B1").Value = Array("Produto", "Vendas")
    ws.Range("A2:B2").Value = Array("Eletroportáteis e utilidades", 50)
    ws.Range("A3:B3").Value = Array("Devoluções", -10)
    ws.Range("A4:B4").Value = Array("Sem venda", 0)
    ws.Range("A5:B5").Value = Array("Indisponível", "n/d")
    ws.Range("A6:B6").Value = Array("Móveis", 30)
    ws.Range("A7:B7").Value = Array("Tintas", 20)
    ws.Range("A8:B8").Value = Array("Ferramentas", 10)
    Set d = ChartData_FromRange(ws.Range("A1:B8"))
    Set cht = Pie_Insert(ws, d, "pie").Chart
    Set s = cht.SeriesCollection(1): Set body = Block_Body(Registry_Block(cht))
    AssertEqual 0, body.Cells(3, 2).Value: AssertEqual 0, body.Cells(5, 2).Value
    AssertEqual -10, ws.Range("B3").Value, "tabela do usuario intacta"
    AssertAlmost 50 / 110, body.Cells(2, 3).Value
    AssertEqual 7, s.Points.Count
    pal = Style_Palette("padrao")
    AssertEqual CLng(pal(3)), s.Points(7).Format.Fill.ForeColor.RGB, "setima fatia: pal(3), nao repete a da 1a (vizinha dela no circulo)"
    AssertTrue s.Points(7).Format.Fill.ForeColor.RGB <> s.Points(1).Format.Fill.ForeColor.RGB, "a ultima fatia difere da primeira"
    AssertEqual "", s.Points(2).DataLabel.Text, "fatia zero sem rotulo"
    Pie_SetOutside cht, True
    ForceLayout cht
    For i = 1 To 7
        Set dl = s.Points(i).DataLabel
        If body.Cells(i + 1, 2).Value > 0 Then
            ' tolerancia de 0,5 pt nas quatro bordas (o Excel encosta o rotulo de cima na borda: medido Top = -0,2)
            AssertTrue dl.Left >= -0.5 And dl.Top >= -0.5, "rotulo " & i & " dentro do grafico (esquerda/topo)"
            AssertTrue dl.Left + dl.Width <= cht.ChartArea.Width + 0.5 And dl.Top + dl.Height <= cht.ChartArea.Height + 0.5, "rotulo " & i & " dentro do grafico (direita/base)"
        Else
            ' fatia sem valor: sem nome na borda (tres nomes de fatias zero se amontoariam no mesmo ponto)
            AssertEqual "", dl.Text, "fatia " & i & " sem valor fica sem rotulo, tambem fora"
        End If
    Next
    ' o nome acompanha a fatia quando os dados mudam (Core_SafeRefresh e o que os eventos chamam)
    ws.Range("B3").Value = 15
    Core_SafeRefresh cht
    AssertTrue InStr(s.Points(2).DataLabel.Text, "Devoluções") > 0 And InStr(s.Points(2).DataLabel.Text, "15") > 0, "fatia que ganhou valor ganha o nome: " & s.Points(2).DataLabel.Text
    ws.Range("B2").Value = 0
    Core_SafeRefresh cht
    AssertEqual "", s.Points(1).DataLabel.Text, "fatia que foi a zero perde o nome"
    Watchers_Stop
    CloseTemp wb
End Sub

Public Sub Test_Pie_Refusals_And_Rollback()
    Dim wb As Workbook, d As clsChartData, cht As Chart, ws As Worksheet, en As Long, ed As String, coll As New Collection, skipped As Long, nm As Name
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Pie_Insert(ws, d, "pie").Chart
    On Error Resume Next
    Ann_AddCagr cht, 1, 3
    en = Err.Number: ed = Err.Description: Err.Clear
    AssertEqual vbObjectError + 7001, en: AssertEqual "A seta CAGR não está disponível para este tipo de gráfico.", ed
    Deco_SetValueLine cht, True
    en = Err.Number: Err.Clear
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    coll.Add cht
    AssertEqual 0, Deco_SameScale(coll, skipped): AssertEqual 1, skipped
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Pie_Insert ws, d, "doughnut"
    en = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertEqual vbObjectError + 9999, en
    AssertEqual 0, ws.ChartObjects.Count
    For Each nm In wb.Names
        AssertTrue Left$(nm.Name, 4) <> "tqq_", "nome orfao " & nm.Name
    Next
    CloseTemp wb
End Sub

Public Sub Test_Core_Insert_Pie_TogglesAndOptions()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:B4").Select
    Core_Insert "pie"
    Set cht = Core_ActiveChart()
    AssertEqual "pie", Registry_ChartType(cht)
    AssertTrue Core_Supports("pieout"): AssertFalse Core_Supports("totals")
    Core_SetToggle "pieout", True
    AssertTrue Core_GetToggle("pieout"): AssertFalse cht.HasLegend
    Core_SetOption "numfmt", "pct1"
    AssertEqual "50" & Application.International(xlDecimalSeparator) & "0%", Mid$(cht.SeriesCollection(1).Points(1).DataLabel.Text, InStr(cht.SeriesCollection(1).Points(1).DataLabel.Text, "50"))
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), cht.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB
    Core_SetToggle "labels", False
    AssertFalse cht.SeriesCollection(1).HasDataLabels
    Core_SetToggle "labels", True
    AssertTrue cht.SeriesCollection(1).DataLabels.ShowPercentage, "religar por Pie_SetLabels mantem pct1 (participacao)"
    AssertTrue cht.SeriesCollection(1).DataLabels.ShowCategoryName, "religar mantem os rotulos fora com o nome"
    AssertEqual xlLabelPositionOutsideEnd, cht.SeriesCollection(1).DataLabels.Position
    Watchers_Stop
    CloseTemp wb
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Testes adicionais (Tarefa 8)
' ---------------------------------------------------------------------------------------------------------------------

Private Function Fixture5(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:B1").Value = Array("Produto", "Vendas")
    ws.Range("A2:B2").Value = Array("Eletro", 50)
    ws.Range("A3:B3").Value = Array("Móveis", 30)
    ws.Range("A4:B4").Value = Array("Tintas", 20)
    ws.Range("A5:B5").Value = Array("Pisos", 15)
    ws.Range("A6:B6").Value = Array("Ferragens", 10)
    Set Fixture5 = ChartData_FromRange(ws.Range("A1:B6"))
End Function

' Uma serie so: a primeira coluna de valores. Direto (Pie_Insert) as outras colunas nao entram; por Core_Insert o usuario e avisado.
Public Sub Test_Pie_ManyValueColumns_UsesFirstAndNotifies()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, v As Variant
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Produto", "Vendas", "Custo")
    ws.Range("A2:C2").Value = Array("Eletro", 50, 900)
    ws.Range("A3:C3").Value = Array("Móveis", 30, 800)
    ws.Range("A4:C4").Value = Array("Tintas", 20, 700)
    Set d = ChartData_FromRange(ws.Range("A1:C4"))
    Set cht = Pie_Insert(ws, d, "pie").Chart
    AssertEqual 1, cht.SeriesCollection.Count, "uma serie so"
    AssertEqual "Vendas", cht.SeriesCollection(1).Name
    v = cht.SeriesCollection(1).Values
    AssertEqual 50, v(1): AssertEqual 20, v(3), "valores da primeira coluna (nao da Custo)"
    AssertEqual 3, cht.SeriesCollection(1).Points.Count
    AssertEqual "50", cht.SeriesCollection(1).Points(1).DataLabel.Text
    ws.ChartObjects(1).Delete
    ' pelo comando: o grafico e criado com a primeira coluna e o usuario e avisado
    ws.Range("A1:C4").Select
    Core_ClearLastNotice
    Core_Insert "pie"
    AssertEqual 1, ws.ChartObjects.Count
    AssertEqual 1, Core_ActiveChart().SeriesCollection.Count
    AssertTrue InStr(Core_LastNotice(), "pizza") > 0 And InStr(Core_LastNotice(), "primeira coluna") > 0 And InStr(Core_LastNotice(), "Vendas") > 0, Core_LastNotice()
    ws.ChartObjects(1).Delete
    Core_ClearLastNotice
    Core_Insert "doughnut"
    AssertTrue InStr(Core_LastNotice(), "rosca") > 0 And InStr(Core_LastNotice(), "primeira coluna") > 0, Core_LastNotice()
    ' uma coluna de valores: sem aviso
    ws.ChartObjects(1).Delete
    ws.Range("A1:B4").Select
    Core_ClearLastNotice
    Core_Insert "pie"
    AssertEqual 0, Core_NoticeCount(), "sem aviso com uma coluna de valores: " & Core_LastNotice()
    Teardown wb
End Sub

' Limpeza de blocos (Tarefa 0) na pizza: a serie cita o bloco (Values), entao uma copia sem o nome tqq_<id> o mantem vivo; apagada, o bloco e
' removido (sem nome orfao).
Public Sub Test_Pie_Cleanup_KeepsBlockReferencedBySeries_AndRemovesOrphan()
    Dim wb As Workbook, d As clsChartData, co As ChartObject, id As String
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set co = Pie_Insert(d.Source.Worksheet, d, "pie")
    id = Registry_IdFromChart(co.Chart)
    AssertEqual 0, CalcSheet_Cleanup(wb), "grafico com o nome tqq_<id>: bloco vivo"
    co.Name = "Copia da pizza"
    AssertEqual 0, CalcSheet_Cleanup(wb), "copia sem o nome: a serie ainda cita o bloco"
    AssertFalse CalcSheet_CleanupReadFailed()
    AssertTrue Not CalcSheet_FindBlock(wb, id) Is Nothing
    AssertEqual 1, TqqNameCount(wb)
    co.Delete
    AssertEqual 1, CalcSheet_Cleanup(wb), "sem grafico: o bloco orfao sai"
    AssertEqual 0, TqqNameCount(wb), "nenhum nome orfao"
    CloseTemp wb
End Sub

' Rollback depois de o grafico existir (gancho Util_FailAfterChartForTest) e tipo fora da familia: nada fica na planilha.
Public Sub Test_Pie_Rollback_AfterChart_AndWrongKind_LeaveNothing()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, k As Variant, en As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    For Each k In Array("pie", "doughnut")
        Util_FailAfterChartForTest = True
        On Error Resume Next
        Pie_Insert ws, d, CStr(k)
        en = Err.Number: Err.Clear
        On Error GoTo 0
        Util_FailAfterChartForTest = False
        AssertEqual vbObjectError + 9998, en, CStr(k)
        AssertEqual 0, ws.ChartObjects.Count, "grafico desfeito (" & k & ")"
        AssertEqual 0, TqqNameCount(wb), "bloco desfeito (" & k & ")"
    Next
    On Error Resume Next
    Pie_Insert ws, d, "stacked"
    en = Err.Number: Err.Clear
    On Error GoTo 0
    AssertEqual 5, en, "tipo fora da familia recusado antes de alocar o bloco"
    AssertEqual 0, TqqNameCount(wb)
    CloseTemp wb
End Sub

' Celula de erro, vazia e tudo zero: o Valor vira 0, sem #N/D nem #DIV/0!, e o grafico nasce.
Public Sub Test_Pie_ErrorEmptyAndAllZero_NoErrorValues()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, body As Range, i As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("B3").Formula = "=NA()"
    ws.Range("B4").ClearContents
    Set cht = Pie_Insert(ws, d, "pie").Chart
    Set body = Block_Body(Registry_Block(cht))
    For i = 2 To 4
        AssertFalse IsError(body.Cells(i, 2).Value), "Valor sem erro na linha " & i
        AssertFalse IsError(body.Cells(i, 3).Value), "Share sem erro na linha " & i
    Next
    AssertEqual 0, body.Cells(3, 2).Value: AssertEqual 0, body.Cells(4, 2).Value
    AssertAlmost 1, body.Cells(2, 3).Value, 0.000001, "so uma fatia positiva: 100%"
    ws.ChartObjects(1).Delete
    ws.Range("B2").Value = 0: ws.Range("B3").Value = 0: ws.Range("B4").Value = 0
    Set cht = Pie_Insert(ws, d, "pie").Chart
    Set body = Block_Body(Registry_Block(cht))
    For i = 2 To 4
        AssertFalse IsError(body.Cells(i, 3).Value), "Share de soma zero sem #DIV/0!"
        AssertEqual 0, body.Cells(i, 3).Value
    Next
    AssertEqual 3, cht.SeriesCollection(1).Points.Count
    CloseTemp wb
End Sub

' Os dados mudam: valores, rotulos e o nome da serie seguem a origem (rotulos nativos sobre a coluna Valor do bloco).
Public Sub Test_Pie_DataChange_UpdatesSeriesAndLabels()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, s As Series, v As Variant
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    Set cht = Pie_Insert(ws, d, "pie").Chart
    Set s = cht.SeriesCollection(1)
    ws.Range("B2").Value = 100
    ws.Range("B1").Value = "Receita"
    v = s.Values
    AssertEqual 100, v(1)
    AssertEqual "100", s.Points(1).DataLabel.Text
    AssertEqual "Receita", s.Name
    Pie_ApplyNumberFormat cht, "pct0"
    AssertEqual "67%", s.Points(1).DataLabel.Text, "100 de 150"
    ws.Range("A2").Value = "Eletrodomésticos"
    Pie_SetOutside cht, True
    AssertTrue InStr(s.Points(1).DataLabel.Text, "Eletrodomésticos") > 0, "categoria nova no rotulo: " & s.Points(1).DataLabel.Text
    CloseTemp wb
End Sub

' Legenda: right/none valem na pizza e na rosca; lineend e recusado (7002) sem mexer na opcao salva nem na legenda.
Public Sub Test_Pie_Legend_RightNone_AndLineEndRefused_BothKinds()
    Dim wb As Workbook, d As clsChartData, cht As Chart, k As Variant, en As Long, ed As String
    For Each k In Array("pie", "doughnut")
        Set wb = NewTempWorkbook(): Set d = Fixture(wb)
        Set cht = Pie_Insert(d.Source.Worksheet, d, CStr(k)).Chart
        AssertTrue cht.HasLegend, CStr(k) & " nasce com legenda"
        Deco_SetLegend cht, "none"
        AssertFalse cht.HasLegend: AssertEqual "none", Registry_GetOption(cht, "legend")
        Deco_SetLegend cht, "right"
        AssertTrue cht.HasLegend: AssertEqual "right", Registry_GetOption(cht, "legend")
        AssertEqual xlLegendPositionRight, cht.Legend.Position
        AssertEqual 3, cht.Legend.LegendEntries.Count
        On Error Resume Next
        Deco_SetLegend cht, "lineend"
        en = Err.Number: ed = Err.Description: Err.Clear
        On Error GoTo 0
        AssertEqual vbObjectError + 7002, en, CStr(k)
        AssertEqual "Rótulos no fim da linha só existem em gráficos de linha e combinação.", ed
        AssertEqual "right", Registry_GetOption(cht, "legend"), "recusa nao grava a opcao"
        AssertTrue cht.HasLegend
        CloseTemp wb
    Next
End Sub

' Paleta: a cor da fonte dos rotulos de dentro segue a fatia nova (expectativas explicitas, sem Style_ContrastFont): a 5a fatia passa do
' verde-azulado escuro (fonte branca) ao cinza claro (fonte cinza-escuro); fora, a fonte e sempre cinza-escuro. Tema: cores accent.
Public Sub Test_Pie_Palette_RecolorsSlicesAndLabelFonts()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, ws As Worksheet
    Set wb = NewTempWorkbook(): Set d = Fixture5(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:B6").Select
    Core_Insert "pie"
    Set cht = Core_ActiveChart(): Set s = cht.SeriesCollection(1)
    AssertEqual RGB(0, 150, 136), s.Points(5).Format.Fill.ForeColor.RGB
    AssertEqual RGB(255, 255, 255), s.Points(5).DataLabel.Font.Color, "fatia escura: fonte branca"
    AssertEqual RGB(64, 64, 64), s.Points(3).DataLabel.Font.Color, "fatia clara: fonte cinza-escuro"
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(217, 217, 217), s.Points(5).Format.Fill.ForeColor.RGB
    AssertEqual RGB(64, 64, 64), s.Points(5).DataLabel.Font.Color, "fatia nova clara: fonte cinza-escuro"
    AssertEqual RGB(255, 255, 255), s.Points(1).DataLabel.Font.Color, "fatia nova escura: fonte branca"
    AssertEqual "cinza", Registry_GetOption(cht, "palette")
    Core_SetOption "palette", "tema"
    ' o Excel devolve a cor de tema como RGB resolvido (ObjectThemeColor le 0): compara com as cores accent do tema do livro
    AssertEqual wb.Theme.ThemeColorScheme.Colors(msoThemeAccent1).RGB, s.Points(1).Format.Fill.ForeColor.RGB, "fatia 1 = accent1"
    AssertEqual wb.Theme.ThemeColorScheme.Colors(msoThemeAccent5).RGB, s.Points(5).Format.Fill.ForeColor.RGB, "fatia 5 = accent5"
    AssertTrue s.Points(1).Format.Fill.ForeColor.RGB <> s.Points(2).Format.Fill.ForeColor.RGB, "cores de tema distintas"
    Core_SetToggle "pieout", True
    AssertEqual RGB(64, 64, 64), s.Points(1).DataLabel.Font.Color, "rotulos fora: sempre cinza-escuro"
    Core_SetToggle "pieout", False
    Core_SetOption "palette", "padrao"
    AssertEqual RGB(255, 255, 255), s.Points(1).DataLabel.Font.Color, "de volta para dentro: contraste com a fatia"
    Teardown wb
End Sub

' Visual base: sem titulo, borda branca de 1 pt entre as fatias e sem eixos (Style_ApplyBase nao pode levantar erro na pizza).
Public Sub Test_Pie_BaseStyle_NoTitleWhiteSliceBorders()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, k As Variant
    For Each k In Array("pie", "doughnut")
        Set wb = NewTempWorkbook(): Set d = Fixture(wb)
        Set cht = Pie_Insert(d.Source.Worksheet, d, CStr(k)).Chart
        Set s = cht.SeriesCollection(1)
        AssertFalse cht.HasTitle, CStr(k)
        AssertEqual msoTrue, s.Format.Line.Visible
        AssertEqual RGB(255, 255, 255), s.Format.Line.ForeColor.RGB
        AssertEqual 1, s.Format.Line.Weight
        AssertTrue cht.ChartGroups(1).VaryByCategories
        Style_ApplyBase cht                                    ' reaplicavel sem erro
        CloseTemp wb
    Next
End Sub

' "Alterar dados" reaplica as ESCOLHAS salvas (rotulos fora, legenda, formato, paleta, rotulos), nao o padrao do Insert nem o estado real do
' grafico novo; a pizza nao ganha opcoes de outros tipos (totals, connectors).
Public Sub Test_Core_RebuildFromRange_Pie_ReappliesSavedOptions()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A5:B5").Value = Array("Pisos", 10)
    ws.Range("A1:B4").Select
    Core_Insert "pie"
    Set cht = Core_ActiveChart()
    Core_SetToggle "pieout", True
    Core_SetOption "numfmt", "pct1"
    Core_SetOption "palette", "cinza"
    AssertTrue cht.HasLegend = False And Registry_GetOption(cht, "legend") = "right", "rotulos fora: legenda oculta, opcao right guardada"
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:B5"))
    Set cht = Core_ActiveChart(): Set s = cht.SeriesCollection(1)
    AssertEqual 1, ws.ChartObjects.Count
    AssertEqual "pie", Registry_ChartType(cht)
    AssertEqual 4, s.Points.Count: AssertEqual "4", Registry_GetOption(cht, "n")
    AssertEqual "1", Registry_GetOption(cht, "pieout")
    AssertEqual xlLabelPositionOutsideEnd, s.DataLabels.Position
    AssertTrue s.DataLabels.ShowCategoryName: AssertTrue s.DataLabels.ShowPercentage
    AssertFalse cht.HasLegend, "rotulos fora vencem a legenda salva (right)": AssertEqual "right", Registry_GetOption(cht, "legend")
    AssertEqual RGB(64, 64, 64), s.Points(1).Format.Fill.ForeColor.RGB
    AssertEqual "", Registry_GetOption(cht, "totals", ""): AssertEqual "", Registry_GetOption(cht, "connectors", "")
    AssertFalse Core_ChartSourceMismatch(cht)
    ' sem rotulos fora a legenda salva (right) volta; com legend=none a legenda continua oculta mesmo apos "Alterar dados"
    Core_SetToggle "pieout", False
    AssertTrue cht.HasLegend
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:B5"))
    Set cht = Core_ActiveChart(): Set s = cht.SeriesCollection(1)
    AssertEqual "0", Registry_GetOption(cht, "pieout")
    AssertEqual xlLabelPositionInsideEnd, s.DataLabels.Position: AssertFalse s.DataLabels.ShowCategoryName
    AssertTrue cht.HasLegend: AssertEqual 4, cht.Legend.LegendEntries.Count
    Core_SetOption "legend", "none"
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:B5"))
    Set cht = Core_ActiveChart()
    AssertFalse cht.HasLegend, "legend=none continua valendo": AssertEqual "none", Registry_GetOption(cht, "legend")
    Core_SetOption "legend", "right"
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:B4"))
    Set cht = Core_ActiveChart(): Set s = cht.SeriesCollection(1)
    AssertTrue cht.HasLegend: AssertEqual 3, cht.Legend.LegendEntries.Count
    AssertEqual "3", Registry_GetOption(cht, "n")
    ' rotulos desligados continuam desligados
    Core_SetToggle "labels", False
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:B4"))
    Set cht = Core_ActiveChart()
    AssertFalse cht.SeriesCollection(1).HasDataLabels: AssertEqual "0", Registry_GetOption(cht, "labels")
    Teardown wb
End Sub

' Recusas pelo Core_*: sem erro ao chamador, aviso com o texto exato e nada gravado.
Public Sub Test_Core_Pie_Refusals_NoticeAndNothingSaved()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, i As Long, n As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:B4").Select
    Core_Insert "doughnut"
    Set cht = Core_ActiveChart()
    AssertFalse Core_Supports("pieout"): AssertTrue Core_Supports("labels"): AssertTrue Core_Supports("numfmt"): AssertTrue Core_Supports("legend")
    AssertFalse Core_Supports("lineend"): AssertFalse Core_Supports("cagr"): AssertFalse Core_Supports("vline"): AssertFalse Core_Supports("samescale")
    Core_ClearLastNotice
    Core_SetToggle "pieout", True
    AssertEqual "'Rótulos fora da pizza' não se aplica a este tipo de gráfico.", Core_LastNotice()
    AssertEqual "", Registry_GetOption(cht, "pieout", ""): AssertTrue cht.HasLegend
    Core_ClearLastNotice
    Core_SetOption "legend", "lineend"
    AssertEqual "Rótulos no fim da linha só existem em gráficos de linha e combinação.", Core_LastNotice()
    AssertEqual "right", Registry_GetOption(cht, "legend"): AssertTrue cht.HasLegend
    Core_ClearLastNotice
    Core_SetToggle "vline", True
    AssertEqual "'Linha de valor (média)' não se aplica a este tipo de gráfico.", Core_LastNotice()
    Core_ClearLastNotice
    Core_AddAnnotation "cagr", 1, 3
    AssertEqual "A seta CAGR não está disponível para este tipo de gráfico.", Core_LastNotice()
    Core_ClearLastNotice
    Core_AddAnnotation "diff", 1, 3
    AssertEqual "A seta de diferença não está disponível para este tipo de gráfico.", Core_LastNotice()
    For i = 1 To cht.Shapes.Count
        If Left$(cht.Shapes(i).Name, 7) = "tqq_ann" Then n = n + 1
    Next
    AssertEqual 0, n, "nenhuma seta criada"
    ' alternadores sem sentido na pizza: ignorados sem gravar opcao de outro tipo
    Core_SetToggle "totals", True
    Core_SetToggle "connectors", True
    AssertEqual "", Registry_GetOption(cht, "totals", ""): AssertEqual "", Registry_GetOption(cht, "connectors", "")
    Teardown wb
End Sub

' O usuario apagou o rotulo de uma fatia: formato, rotulos fora e atualizar seguem sem erro ao chamador (e religar os rotulos recria todos).
Public Sub Test_Pie_UserDeletedOneLabel_OptionsStillApply()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:B4").Select
    Core_Insert "pie"
    Set cht = Core_ActiveChart(): Set s = cht.SeriesCollection(1)
    Core_SetToggle "pieout", True
    s.Points(2).HasDataLabel = False
    Core_ClearLastNotice
    Core_SetOption "numfmt", "pct0"
    AssertTrue InStr(s.Points(1).DataLabel.Text, "Eletro") > 0 And InStr(s.Points(1).DataLabel.Text, "%") > 0, s.Points(1).DataLabel.Text
    s.Points(2).HasDataLabel = False
    ws.Range("B4").Value = 0                                ' a fatia 3 vai a zero depois do rotulo da 2 ter sido apagado
    Core_SafeRefresh cht
    AssertEqual 0, Core_NoticeCount(), "sem aviso de erro: " & Core_LastNotice()
    AssertEqual "", s.Points(3).DataLabel.Text, "a fatia que foi a zero perde o nome mesmo com o rotulo vizinho apagado"
    AssertFalse s.Points(2).HasDataLabel, "atualizar nao recria o rotulo que o usuario apagou"
    Core_SetToggle "labels", False
    Core_SetToggle "labels", True
    AssertTrue s.Points(2).HasDataLabel And s.Points(3).HasDataLabel, "religar recria todos os rotulos"
    AssertEqual 0, Core_NoticeCount(), Core_LastNotice()
    Teardown wb
End Sub

' Geometria (regra de tres no PlotArea real, rotulos lidos depois de renderizar): na pizza o rotulo de dentro fica perto da borda (InsideEnd:
' entre 55% e 100% do raio); na rosca, sobre o anel (entre o furo, 50% do raio, e a borda); as linhas de chamada so existem com rotulos fora.
Public Sub Test_Pie_LabelGeometry_InsideNearEdge_DoughnutOnRing_LeaderLines()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series, k As Long, dd As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Pie_Insert(d.Source.Worksheet, d, "pie").Chart: Set s = cht.SeriesCollection(1)
    ForceLayout cht
    For k = 1 To 3
        dd = Dist(cht, s.Points(k).DataLabel)
        AssertTrue dd > 0.55 * Radius(cht) And dd < Radius(cht), "pizza: rotulo " & k & " perto da borda, dentro do circulo (" & dd & " de " & Radius(cht) & ")"
    Next
    AssertFalse s.HasLeaderLines
    Pie_SetOutside cht, True
    AssertTrue s.HasLeaderLines, "rotulos fora: linhas de chamada"
    Pie_SetOutside cht, False
    AssertFalse s.HasLeaderLines
    CloseTemp wb
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Pie_Insert(d.Source.Worksheet, d, "doughnut").Chart: Set s = cht.SeriesCollection(1)
    ForceLayout cht
    For k = 1 To 3
        dd = Dist(cht, s.Points(k).DataLabel)
        AssertTrue dd > 0.5 * Radius(cht) And dd < Radius(cht), "rosca: rotulo " & k & " sobre o anel (" & dd & " de " & Radius(cht) & ")"
    Next
    CloseTemp wb
End Sub

' Atualizar (Core_SafeRefresh) nao mexe nos rotulos nativos nem cria formas na pizza: a posicao (dentro/fora) e o conteudo ficam.
Public Sub Test_Pie_SafeRefresh_KeepsLabelsAndCreatesNoShapes()
    Dim wb As Workbook, d As clsChartData, cht As Chart, s As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Pie_Insert(d.Source.Worksheet, d, "pie").Chart
    Set s = cht.SeriesCollection(1)
    Core_SafeRefresh cht: Labels_RefreshPositions cht: Ann_Refresh cht
    Labels_ApplyCellFormat cht, Style_NumberFormat("dec2")       ' os rotulos sao nativos: nenhuma celula do bloco recebe formato
    AssertEqual "General", Block_Body(Registry_Block(cht)).Cells(2, 2).NumberFormat
    AssertEqual xlLabelPositionInsideEnd, s.DataLabels.Position
    AssertEqual "50", s.Points(1).DataLabel.Text
    Pie_SetOutside cht, True
    Core_SafeRefresh cht
    AssertEqual xlLabelPositionOutsideEnd, s.DataLabels.Position
    AssertTrue s.DataLabels.ShowCategoryName
    AssertEqual 0, cht.Shapes.Count, "sem anotacoes"
    AssertFalse Core_ChartSourceMismatch(cht), "origem do tamanho esperado"
    Watchers_Stop
    CloseTemp wb
End Sub

' Tarefa 11 (M1 da Tarefa 8): rotulos fora SEM rotulos deixavam o grafico sem legenda e sem nomes. A legenda so some quando os nomes estao
' nos rotulos: (legend <> "none") And Not (pieout = "1" And labels <> "0"). Vale nas duas ordens, com legend=none e depois de "Alterar dados".
Public Sub Test_Pie_OutsideWithLabelsOff_KeepsLegend()
    Dim wb As Workbook, d As clsChartData, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb): Set ws = d.Source.Worksheet
    ws.Range("A1:B4").Select
    Core_Insert "pie"
    Set cht = Core_ActiveChart()
    Core_SetToggle "pieout", True
    AssertFalse cht.HasLegend, "rotulos fora ligados: os nomes estao nos rotulos, sem legenda"
    Core_SetToggle "labels", False
    AssertTrue cht.HasLegend, "rotulos desligados: a legenda volta (senao nao haveria nome em lugar nenhum)"
    AssertEqual 3, cht.Legend.LegendEntries.Count
    AssertEqual "1", Registry_GetOption(cht, "pieout"): AssertEqual "right", Registry_GetOption(cht, "legend"), "a opcao de legenda nao e reescrita"
    Core_SetToggle "labels", True
    AssertFalse cht.HasLegend, "rotulos religados: a legenda sai de novo"
    ' outra ordem: rotulos desligados primeiro, depois os rotulos fora
    Core_SetToggle "pieout", False
    Core_SetToggle "labels", False
    AssertTrue cht.HasLegend
    Core_SetToggle "pieout", True
    AssertTrue cht.HasLegend, "pieout ligado com rotulos desligados mantem a legenda"
    ' legend=none vence sempre
    Core_SetOption "legend", "none"
    AssertFalse cht.HasLegend
    Core_SetOption "legend", "right"
    AssertTrue cht.HasLegend
    ' Alterar dados com pieout=1 e labels=0: a legenda continua
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:B4"))
    Set cht = Core_ActiveChart()
    AssertEqual "1", Registry_GetOption(cht, "pieout"): AssertEqual "0", Registry_GetOption(cht, "labels")
    AssertTrue cht.HasLegend, "apos Alterar dados: pieout=1 e labels=0 mantem a legenda"
    AssertFalse cht.SeriesCollection(1).HasDataLabels
    Teardown wb
End Sub

' Tarefa 11 (M2 da Tarefa 8): com n > 6 e (n - 1) Mod 6 = 0 a ultima fatia cairia na cor da primeira (e as duas sao vizinhas no circulo):
' usa pal(3). Nos demais n o ciclo e o de sempre (a fatia k recebe pal((k - 1) Mod 6)). Expectativas escritas a mao, sem a formula de producao.
Public Sub Test_Pie_LastSlice_DiffersFromFirst_WhenCycleWraps()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, s As Series, pal As Variant, n As Variant, i As Long, want As Variant
    Dim idx13 As Variant, idx8 As Variant
    idx13 = Array(0, 1, 2, 3, 4, 5, 0, 1, 2, 3, 4, 5, 3)      ' 13 fatias: a 13a (que seria pal(0)) vira pal(3)
    idx8 = Array(0, 1, 2, 3, 4, 5, 0, 1)                      ' 8 fatias: o ciclo normal (a 8a e pal(1), vizinha da 1a que e pal(0))
    pal = Style_Palette("padrao")
    For Each n In Array(13, 8)
        Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
        ws.Range("A1:B1").Value = Array("Produto", "Vendas")
        For i = 1 To n
            ws.Cells(i + 1, 1).Value = "P" & i: ws.Cells(i + 1, 2).Value = 10 + i
        Next
        Set d = ChartData_FromRange(ws.Range("A1").Resize(n + 1, 2))
        Set cht = Pie_Insert(ws, d, IIf(n = 13, "pie", "doughnut")).Chart
        Set s = cht.SeriesCollection(1)
        If n = 13 Then want = idx13 Else want = idx8
        AssertEqual n, s.Points.Count
        For i = 1 To n
            AssertEqual CLng(pal(want(i - 1))), s.Points(i).Format.Fill.ForeColor.RGB, n & " fatias: fatia " & i
        Next
        AssertTrue s.Points(n).Format.Fill.ForeColor.RGB <> s.Points(1).Format.Fill.ForeColor.RGB, n & " fatias: a ultima difere da primeira"
        ' a regra vale tambem para a paleta trocada pela API
        Core_ClearLastNotice
        cht.Parent.Select
        Core_SetOption "palette", "cinza"
        pal = Style_Palette("cinza")
        AssertEqual CLng(pal(want(n - 1))), s.Points(n).Format.Fill.ForeColor.RGB, n & " fatias: ultima fatia na paleta cinza"
        AssertTrue s.Points(n).Format.Fill.ForeColor.RGB <> s.Points(1).Format.Fill.ForeColor.RGB, n & " fatias: cinza, a ultima difere da primeira"
        pal = Style_Palette("padrao")
        Teardown wb
    Next
End Sub
