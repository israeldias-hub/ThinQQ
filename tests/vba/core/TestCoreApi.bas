Attribute VB_Name = "TestCoreApi"
Option Explicit

Private Function Fixture(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 100, 50)
    ws.Range("A3:C3").Value = Array(2023, 110, 60)
    ws.Range("A4:C4").Value = Array(2024, 120, 70)
    ws.Range("A5:C5").Value = Array(2025, 150, 90)
    Set Fixture = ws
End Function

Public Sub Test_Core_HasChart_FalseOnCells()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1").Select
    AssertFalse Core_HasChart()
    AssertTrue Core_ActiveChart() Is Nothing
    Teardown wb
End Sub

Public Sub Test_Core_Insert_SelectsChart_And_Hooks()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_OnLoaded
    Core_Insert "stacked"
    AssertTrue Core_HasChart(), "o gráfico inserido deve ficar selecionado"
    AssertEqual "stacked", Registry_ChartType(Core_ActiveChart())
    AssertEqual 1, Watchers_Count()
    AssertTrue Core_GetToggle("totals")
    AssertTrue Core_GetToggle("labels")
    AssertFalse Core_GetToggle("connectors")
    AssertEqual "int", Core_GetOption("numfmt")
    Core_OnUnload
    Teardown wb
End Sub

Public Sub Test_Core_Toggles_And_Options()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetToggle "totals", False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertFalse Core_GetToggle("totals")
    Core_SetToggle "connectors", True
    AssertTrue cht.ChartGroups(1).HasSeriesLines
    Core_SetOption "numfmt", "dec1"
    AssertEqual "100" & Application.International(xlDecimalSeparator) & "0", cht.SeriesCollection(1).Points(1).DataLabel.Text
    Core_SetOption "gap", "120"
    AssertEqual 120, cht.ChartGroups(1).GapWidth
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), cht.SeriesCollection(1).Format.Fill.ForeColor.RGB
    Teardown wb
End Sub

Public Sub Test_Core_Insert_Waterfall_And_Annotations()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Início", 100)
    ws.Range("A3:B3").Value = Array("A", -30)
    ws.Range("A4:B4").Value = Array("Fim", "e")
    ws.Range("A1").Select
    Core_Insert "waterfall"
    Set cht = Core_ActiveChart()
    AssertEqual "waterfall", Registry_ChartType(cht)
    AssertTrue Core_GetToggle("connectors")
    Core_AddAnnotation "cagr", 1, 3
    Core_AddAnnotation "diff", 1, 3
    AssertTrue Not cht.Shapes("tqq_ann_cagr_1") Is Nothing
    AssertTrue Not cht.Shapes("tqq_ann_diff_1") Is Nothing
    Core_RemoveAnnotations
    AssertEqual 2, cht.Shapes.Count, "só os 2 conectores devem restar"
    Core_SetToggle "connectors", False
    AssertEqual 0, cht.Shapes.Count
    Teardown wb
End Sub

Public Sub Test_Core_Insert_InvalidSelection_NoChart()
    Dim wb As Workbook, ws As Worksheet, before As Long, prevAlerts As Boolean
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    before = ws.ChartObjects.Count
    ws.Range("A1:A5").Select
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Core_Insert "stacked"       ' mostra MsgBox apenas quando interativo; em testes Core_Insert não deve travar
    Application.DisplayAlerts = prevAlerts
    AssertEqual before, ws.ChartObjects.Count
    Teardown wb
End Sub

Public Sub Test_Core_Refresh_CleansOrphans()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    ws.ChartObjects(1).Delete
    AssertEqual 1, TqqNameCount(wb), "bloco existe antes da limpeza"
    Core_Refresh
    AssertEqual 0, TqqNameCount(wb)
    Teardown wb
End Sub

Public Sub Test_Core_AddAnnotation_Diff_On100Chart_NoThrow()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked100"
    Set cht = Core_ActiveChart()
    AssertTrue Not cht Is Nothing
    Core_AddAnnotation "diff", 1, 3
    AssertEqual 0, cht.Shapes.Count, "seta de diferença não deve ser criada em 100%"
    Teardown wb
End Sub

Public Sub Test_Core_RebuildFromRange_KeepsPositionOptionsToggles()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, nc As Chart, l As Double, t As Double
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("E1:F1").Value = Array("Ano", "Leste")
    ws.Range("E2:F2").Value = Array(2022, 5)
    ws.Range("E3:F3").Value = Array(2023, 6)
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetToggle "totals", False
    Core_SetToggle "connectors", True
    Core_SetOption "numfmt", "dec1"
    Core_SetOption "gap", "120"
    l = ws.ChartObjects(1).Left: t = ws.ChartObjects(1).Top
    AssertTrue Core_RebuildFromRange(cht, ws.Range("E1:F3"))
    AssertEqual 1, ws.ChartObjects.Count
    Set nc = Core_ActiveChart()
    AssertTrue Not nc Is Nothing
    AssertEqual "stacked", Registry_ChartType(nc)
    AssertEqual l, ws.ChartObjects(1).Left
    AssertEqual t, ws.ChartObjects(1).Top
    AssertEqual "dec1", Registry_GetOption(nc, "numfmt", "")
    AssertFalse Core_GetToggle("totals")
    AssertTrue Core_GetToggle("connectors")
    AssertTrue Chart_FindSeries(nc, "tqq_total") Is Nothing
    AssertTrue nc.ChartGroups(1).HasSeriesLines
    AssertEqual 120, nc.ChartGroups(1).GapWidth
    AssertEqual 2, Geom_ChartCategoryCount(nc)
    AssertEqual 1, TqqNameCount(wb), "bloco antigo deve ser removido"
    Teardown wb
End Sub

Public Sub Test_Watchers_Rehook_CountsOnlyActiveSheetCharts()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    AssertEqual 1, Watchers_Count()
    Watchers_Stop
    AssertEqual 0, Watchers_Count()
    Watchers_Rehook
    AssertEqual 1, Watchers_Count()
    Watchers_Stop
    Teardown wb
End Sub

Public Sub Test_frmAnnotation_PrepareAndResult()
    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmAnnotation")
    frm.Prepare Array("2022", "2023", "2024"), 1, 3
    AssertEqual "cancel", frm.Result
    AssertEqual 0, frm.cboFrom.ListIndex
    AssertEqual 2, frm.cboTo.ListIndex
    AssertEqual 3, frm.cboTo.ListCount
    Unload frm
End Sub

' Conta só os nomes definidos do ThinQQ (o Excel pode manter outros nomes no livro).
Private Function TqqNameCount(wb As Workbook) As Long
    Dim nm As Name
    For Each nm In wb.Names
        If Left$(nm.Name, Len(ID_PREFIX)) = ID_PREFIX Then TqqNameCount = TqqNameCount + 1
    Next
End Function

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

' Monta "<a><milhar><b><decimal><c>" com os separadores do Excel (independe da lógica de produção).
Private Function Sep(a As String, b As String, c As String) As String
    Sep = a & Application.International(xlThousandsSeparator) & b & Application.International(xlDecimalSeparator) & c
End Function

' Insere um empilhado com 2 categorias (valores v1, v2), aplica numfmt=key e devolve o texto exibido do rótulo
' do ponto 1 da série 1 (which=1) ou do ponto 1 dos totais (which=2) ou do ponto 2 da série 1 (which=3).
Private Function Rendered(key As String, v1 As Double, v2 As Double, which As Long, wb As Workbook) As String
    Dim ws As Worksheet, cht As Chart
    Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Ano", "Valor")
    ws.Range("A2").Value = 2022: ws.Range("B2").Value = v1
    ws.Range("A3").Value = 2023: ws.Range("B3").Value = v2
    ws.Range("A1:B3").Select
    Core_Insert "stacked"
    Core_SetOption "numfmt", key
    Set cht = Core_ActiveChart()
    Select Case which
        Case 1: Rendered = cht.SeriesCollection(1).Points(1).DataLabel.Text
        Case 2: Rendered = Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
        Case Else: Rendered = cht.SeriesCollection(1).Points(2).DataLabel.Text
    End Select
End Function

Public Sub Test_Core_NumFmt_RenderedText()
    Dim wb As Workbook, i As Long
    For i = 1 To 2
        Set wb = NewTempWorkbook()
        AssertEqual Sep("1", "234", "5"), Rendered("dec1", 1234.5, 1, i, wb), "dec1 which=" & i
        Teardown wb
        Set wb = NewTempWorkbook()
        AssertEqual "1" & Application.International(xlThousandsSeparator) & "234", Rendered("int", 1234, 1, i, wb), "int which=" & i
        Teardown wb
        Set wb = NewTempWorkbook()
        AssertEqual "-1" & Application.International(xlThousandsSeparator) & "234", Rendered("int", -1234, 1, i, wb), "int negativo which=" & i
        Teardown wb
        Set wb = NewTempWorkbook()
        AssertEqual "12" & Application.International(xlDecimalSeparator) & "3%", Rendered("pct1", 0.123, 1, i, wb), "pct1 which=" & i
        Teardown wb
        Set wb = NewTempWorkbook()
        AssertEqual "1" & Application.International(xlThousandsSeparator) & "235", Rendered("mil", 1234567, 1, i, wb), "mil which=" & i
        Teardown wb
    Next
End Sub

Public Sub Test_Core_NumFmt_ZeroRendersEmpty()
    Dim wb As Workbook
    Set wb = NewTempWorkbook()
    AssertEqual "", Rendered("int", 5, 0, 3, wb), "zero oculto em int"
    Teardown wb
    Set wb = NewTempWorkbook()
    AssertEqual "", Rendered("dec1", 5, 0, 3, wb), "zero oculto em dec1"
    Teardown wb
End Sub

Public Sub Test_Core_InvalidKeysAndValues_DoNotRaise()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, gap As Long
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetToggle "chave_invalida", True
    AssertEqual "", Core_GetOption("chave_invalida")
    gap = cht.ChartGroups(1).GapWidth
    Core_SetOption "gap", "abc"
    Core_SetOption "gap", "9999"
    AssertEqual gap, cht.ChartGroups(1).GapWidth
    AssertEqual "50", Core_GetOption("gap"), "valor inválido não é persistido"
    Core_SetOption "numfmt", "xyz"
    AssertEqual "int", Core_GetOption("numfmt")
    Core_SetOption "palette", "rosa"
    AssertEqual "padrao", Core_GetOption("palette")
    Teardown wb
End Sub

Public Sub Test_Core_Insert_UnknownKind_NoChart()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "pizza"
    AssertEqual 0, ws.ChartObjects.Count
    Teardown wb
End Sub

Public Sub Test_Core_Insert_ProtectedSheet_NoChart()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    ws.Protect
    Core_Insert "stacked"
    ws.Unprotect
    AssertEqual 0, ws.ChartObjects.Count
    AssertEqual 0, TqqNameCount(wb)
    Teardown wb
End Sub

Public Sub Test_Core_Insert_RestoresScreenUpdating()
    Dim wb As Workbook, ws As Worksheet, prev As Boolean
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    prev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Core_Insert "stacked"
    AssertFalse Application.ScreenUpdating, "mantém o valor anterior (False)"
    Application.ScreenUpdating = prev
    Teardown wb
End Sub

Public Sub Test_Core_DeleteChart_ThenInsert_RemovesOrphan()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    ws.ChartObjects(1).Delete
    AssertEqual 1, TqqNameCount(wb)
    ws.Range("A1:C5").Select
    Core_Insert "clustered"
    AssertEqual 1, ws.ChartObjects.Count
    AssertEqual 1, TqqNameCount(wb), "só o bloco do novo gráfico"
    Teardown wb
End Sub

Public Sub Test_Core_RebuildFromRange_Failure_KeepsOriginal()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, id As String, ok As Boolean
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("E1:F1").Value = Array("Ano", "Leste")
    ws.Range("E2:F2").Value = Array(2022, 5)
    ws.Range("E3:F3").Value = Array(2023, 6)
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    id = Registry_IdFromChart(cht)
    Util_FailAfterBlockForTest = True
    ok = Core_RebuildFromRange(cht, ws.Range("E1:F3"))
    Util_FailAfterBlockForTest = False
    AssertFalse ok
    AssertEqual 1, ws.ChartObjects.Count
    AssertEqual ID_PREFIX & id, ws.ChartObjects(1).Name
    AssertTrue Not CalcSheet_FindBlock(wb, id) Is Nothing, "bloco original intacto"
    AssertEqual 1, TqqNameCount(wb), "sem bloco órfão do novo"
    Teardown wb
End Sub

Public Sub Test_Core_Refresh_NoActiveChart_NoThrow()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1").Select
    Core_Refresh
    Teardown wb
End Sub
