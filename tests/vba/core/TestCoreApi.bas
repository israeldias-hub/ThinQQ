Attribute VB_Name = "TestCoreApi"
Option Explicit
Private mSavedPending As String
Private mSavedPendingSha As String
Private mSavedLoaderPath As String
Private mSavedSwap As String
Private mSavedFailNotified As String

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

' Formas de seta de diferenca (tqq_ann_diff*): o 100% tem tambem o indicador 100% (tqq_ann_ind100*), que nao conta.
Private Function CountDiffShapes(cht As Chart) As Long
    Dim i As Long
    For i = 1 To cht.Shapes.Count
        If Left$(cht.Shapes(i).Name, Len("tqq_ann_diff")) = "tqq_ann_diff" Then CountDiffShapes = CountDiffShapes + 1
    Next
End Function

Public Sub Test_Core_AddAnnotation_Diff_On100Chart_NoThrow()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked100"
    Set cht = Core_ActiveChart()
    AssertTrue Not cht Is Nothing
    Core_AddAnnotation "diff", 1, 3
    AssertEqual 0, CountDiffShapes(cht), "seta de diferença não deve ser criada em 100%"
    AssertEqual 2, cht.Shapes.Count, "o 100% so tem o indicador (tique e rotulo)"
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
    AssertEqual "2", Registry_GetOption(nc, "n", ""), "n acompanha a nova tabela (as opcoes antigas trazem o n da anterior)"
    AssertFalse Core_ChartSourceMismatch(nc), "sem divergencia falsa depois de reconstruir com outro tamanho"
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
    Core_Insert "xyz"
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

Public Sub Test_Core_NumFmt_ZeroOnLinkedTotal_RendersEmpty()
    Dim wb As Workbook
    Set wb = NewTempWorkbook()
    AssertEqual "", Rendered("int", 0, 5, 2, wb), "total zero oculto (int)"
    Teardown wb
    Set wb = NewTempWorkbook()
    AssertEqual "", Rendered("dec1", 0, 5, 2, wb), "total zero oculto (dec1)"
    Teardown wb
End Sub

Public Sub Test_Core_NumFmt_Waterfall_RenderedText()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, s As Series, th As String, dc As String, k As Long
    Dim keys As Variant
    th = Application.International(xlThousandsSeparator): dc = Application.International(xlDecimalSeparator)
    For k = 0 To 1
        Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
        ws.Range("A1:B1").Value = Array("Etapa", "Valor")
        ws.Range("A2:B2").Value = Array("Início", 1000)
        ws.Range("A3:B3").Value = Array("A", -150)
        ws.Range("A4:B4").Value = Array("Fim", "e")
        ws.Range("A1").Select
        Core_Insert "waterfall"
        Set cht = Core_ActiveChart()
        Set s = Chart_FindSeries(cht, "tqq_lbl")
        AssertTrue Not s Is Nothing
        If k = 0 Then
            Core_SetOption "numfmt", "int"
            AssertEqual "1" & th & "000", s.Points(1).DataLabel.Text, "cascata int inicio"
            AssertEqual "-150", s.Points(2).DataLabel.Text, "cascata int delta negativo"
            AssertEqual "850", s.Points(3).DataLabel.Text, "cascata int fim"
        Else
            Core_SetOption "numfmt", "dec1"
            AssertEqual "1" & th & "000" & dc & "0", s.Points(1).DataLabel.Text, "cascata dec1 inicio"
            AssertEqual "-150" & dc & "0", s.Points(2).DataLabel.Text, "cascata dec1 delta negativo"
        End If
        Teardown wb
    Next
End Sub

Public Sub Test_Core_SetOption_Gap_OnLineChart_NotPersisted()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "line"
    AssertEqual "line", Registry_ChartType(Core_ActiveChart())
    Core_SetOption "gap", "80"
    AssertEqual "", Core_GetOption("gap"), "gap não se aplica a linha e não é persistido"
    Teardown wb
End Sub

Public Sub Test_Core_RebuildFromRange_SourceOnOtherSheet_KeepsSheetAndPosition()
    Dim wb As Workbook, ws As Worksheet, ws2 As Worksheet, cht As Chart, nc As Chart, l As Double, t As Double, w As Double, h As Double
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set ws2 = wb.Worksheets.Add(After:=ws)
    ws2.Range("A1:B1").Value = Array("Ano", "Leste")
    ws2.Range("A2:B2").Value = Array(2022, 5)
    ws2.Range("A3:B3").Value = Array(2023, 6)
    ws.Activate
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetOption "numfmt", "dec1"
    l = ws.ChartObjects(1).Left: t = ws.ChartObjects(1).Top: w = ws.ChartObjects(1).Width: h = ws.ChartObjects(1).Height
    AssertTrue Core_RebuildFromRange(cht, ws2.Range("A1:B3"))
    AssertEqual 1, ws.ChartObjects.Count, "novo gráfico na planilha do antigo"
    AssertEqual 0, ws2.ChartObjects.Count
    AssertEqual l, ws.ChartObjects(1).Left
    AssertEqual t, ws.ChartObjects(1).Top
    AssertEqual w, ws.ChartObjects(1).Width
    AssertEqual h, ws.ChartObjects(1).Height
    Set nc = ws.ChartObjects(1).Chart
    AssertEqual "dec1", Registry_GetOption(nc, "numfmt", "")
    AssertEqual 2, Geom_ChartCategoryCount(nc)
    AssertEqual 1, TqqNameCount(wb)
    Teardown wb
End Sub

Public Sub Test_Core_RebuildFromRange_ChartSheetNotActive_NewChartSurvives()
    Dim wb As Workbook, ws As Worksheet, ws2 As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set ws2 = wb.Worksheets.Add(After:=ws)
    ws.Activate
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    ws2.Activate        ' o co.Select final falharia se a planilha do gráfico não fosse reativada
    AssertTrue Core_RebuildFromRange(cht, ws.Range("A1:B4"))
    AssertEqual 1, ws.ChartObjects.Count, "o novo gráfico sobrevive"
    AssertEqual 1, TqqNameCount(wb)
    Teardown wb
End Sub


Private Sub PendingSetup()
    mSavedPending = GetSetting("ThinQQ", "Updater", "PendingLoaderPath", "")
    mSavedPendingSha = GetSetting("ThinQQ", "Updater", "PendingLoaderSha256", "")
    mSavedLoaderPath = Core_LoaderPath()
    mSavedSwap = GetSetting("ThinQQ", "Updater", "SwapInProgress", "")
    mSavedFailNotified = GetSetting("ThinQQ", "Updater", "PendingLoaderFailNotified", "")
    PutSetting "SwapInProgress", ""
    PutSetting "PendingLoaderFailNotified", ""
End Sub

Private Sub PutSetting(key As String, value As String)
    If Len(value) > 0 Then
        SaveSetting "ThinQQ", "Updater", key, value
    Else
        On Error Resume Next
        DeleteSetting "ThinQQ", "Updater", key
        On Error GoTo 0
    End If
End Sub

Private Sub PendingRestore()
    On Error Resume Next
    Application.Workbooks("ThinQQ.xlam").Close SaveChanges:=False
    PutSetting "PendingLoaderPath", mSavedPending
    PutSetting "PendingLoaderSha256", mSavedPendingSha
    PutSetting "SwapInProgress", mSavedSwap
    PutSetting "PendingLoaderFailNotified", mSavedFailNotified
    Core_TestSetApplyBusy False
    Core_RegisterLoader "", mSavedLoaderPath
End Sub

Private Sub PendingFail()
    Dim n As Long, s As String, ds As String
    n = Err.Number: s = Err.Source: ds = Err.Description
    Application.DisplayAlerts = True
    PendingRestore
    Err.Raise n, s, ds
End Sub

Private Sub WriteText(path As String, text As String)
    Dim f As Integer
    f = FreeFile: Open path For Output As #f: Print #f, text: Close #f
End Sub

' .xlam valido e vazio (sem macros) gerado no caminho informado.
Private Sub MakeAddinFile(path As String)
    Dim wb As Workbook, prev As Boolean
    prev = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Set wb = Application.Workbooks.Add
    wb.SaveAs Filename:=path, FileFormat:=55      ' xlOpenXMLAddIn
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = prev
End Sub

' Core_ApplyPendingLoader contra uma copia temporaria (nunca o loader real em execucao).
Public Sub Test_Core_ApplyPendingLoader_ReplacesTargetAndClearsSetting()
    Dim d As String, pend As String, target As String, pendLen As Long
    PendingSetup
    On Error GoTo Fail
    d = TempDir()
    MkDir d & "\pend": MkDir d & "\inst"
    pend = d & "\pend\ThinQQ.xlam": target = d & "\inst\ThinQQ.xlam"
    MakeAddinFile pend
    pendLen = FileLen(pend)
    WriteText target, "conteudo antigo do loader"
    AssertTrue FileLen(target) <> pendLen, "pré-condição: alvo difere do pendente"
    Core_RegisterLoader "", target
    SaveSetting "ThinQQ", "Updater", "PendingLoaderPath", pend
    SaveSetting "ThinQQ", "Updater", "PendingLoaderSha256", UCase$(Util_Sha256File(pend))     ' maiusculas: comparacao case-insensitive
    Core_ApplyPendingLoader
    AssertEqual pendLen, FileLen(target), "alvo deve ter o conteúdo do pendente"
    AssertEqual 0, Len(Dir(pend)), "pendente deve ser apagado"
    AssertEqual 0, Len(Dir(target & ".new")) + Len(Dir(target & ".bak")), "sem sobras .new/.bak"
    AssertEqual "", GetSetting("ThinQQ", "Updater", "PendingLoaderPath", ""), "configuração deve ser limpa"
    AssertEqual "", GetSetting("ThinQQ", "Updater", "PendingLoaderSha256", "")
    AssertEqual "", GetSetting("ThinQQ", "Updater", "SwapInProgress", ""), "flag de troca limpa no sucesso"
    PendingRestore
    Exit Sub
Fail:
    PendingFail
End Sub

Public Sub Test_Core_ApplyPendingLoader_ShaMismatch_DiscardsAndKeepsTarget()
    Dim d As String, pend As String, target As String, oldLen As Long
    PendingSetup
    On Error GoTo Fail
    d = TempDir()
    pend = d & "\pend.xlam": target = d & "\ThinQQ.xlam"
    MakeAddinFile pend
    WriteText target, "intacto"
    oldLen = FileLen(target)
    Core_RegisterLoader "", target
    SaveSetting "ThinQQ", "Updater", "PendingLoaderPath", pend
    SaveSetting "ThinQQ", "Updater", "PendingLoaderSha256", "0000000000000000000000000000000000000000000000000000000000000000"
    Core_ApplyPendingLoader
    AssertEqual oldLen, FileLen(target), "alvo intacto"
    AssertEqual 0, Len(Dir(pend)), "arquivo adulterado deve ser apagado"
    AssertEqual "", GetSetting("ThinQQ", "Updater", "PendingLoaderPath", "")
    AssertEqual "", GetSetting("ThinQQ", "Updater", "PendingLoaderSha256", "")
    ' sem SHA registrado: tambem recusa
    MakeAddinFile pend
    SaveSetting "ThinQQ", "Updater", "PendingLoaderPath", pend
    PutSetting "PendingLoaderSha256", ""
    Core_ApplyPendingLoader
    AssertEqual oldLen, FileLen(target)
    AssertEqual 0, Len(Dir(pend))
    PendingRestore
    Exit Sub
Fail:
    PendingFail
End Sub

Public Sub Test_Core_ApplyPendingLoader_ReloadFails_RestoresOldAndKeepsPending()
    Dim d As String, pend As String, target As String
    PendingSetup
    On Error GoTo Fail
    d = TempDir()
    pend = d & "\pend.xlam": target = d & "\ThinQQ.xlam"
    WriteText pend, "isto nao e um xlam valido"
    WriteText target, "loader antigo"
    Core_RegisterLoader "", target
    SaveSetting "ThinQQ", "Updater", "PendingLoaderPath", pend
    SaveSetting "ThinQQ", "Updater", "PendingLoaderSha256", Util_Sha256File(pend)
    Core_ApplyPendingLoader
    AssertEqual "loader antigo", Replace(Replace(ReadText(target), vbCr, ""), vbLf, ""), "loader antigo restaurado"
    AssertTrue Len(Dir(pend)) > 0, "pendente mantido"
    AssertEqual pend, GetSetting("ThinQQ", "Updater", "PendingLoaderPath", ""), "pendente continua registrado"
    AssertEqual 0, Len(Dir(target & ".new")) + Len(Dir(target & ".bak")), "sem sobras .new/.bak"
    AssertEqual "", GetSetting("ThinQQ", "Updater", "SwapInProgress", ""), "flag de troca limpa na falha"
    AssertTrue Len(GetSetting("ThinQQ", "Updater", "PendingLoaderFailNotified", "")) > 0, "falha registrada para o aviso 1x/24 h"
    ' segunda falha em menos de 24 h: o carimbo do aviso nao muda
    SaveSetting "ThinQQ", "Updater", "PendingLoaderFailNotified", Format$(Now - 1 / 24, "yyyy-mm-dd hh:nn:ss")
    Dim stamp As String
    stamp = GetSetting("ThinQQ", "Updater", "PendingLoaderFailNotified", "")
    Core_ApplyPendingLoader
    AssertEqual stamp, GetSetting("ThinQQ", "Updater", "PendingLoaderFailNotified", ""), "aviso limitado a 1x por 24 h"
    PendingRestore
    Exit Sub
Fail:
    PendingFail
End Sub

Public Sub Test_Core_ApplyPendingLoader_MissingPendingFile_ClearsSettingAndKeepsTarget()
    Dim d As String, target As String
    PendingSetup
    On Error GoTo Fail
    d = TempDir()
    target = d & "\ThinQQ.xlam"
    WriteText target, "intacto"
    Core_RegisterLoader "", target
    SaveSetting "ThinQQ", "Updater", "PendingLoaderPath", d & "\nao-existe.xlam"
    Core_ApplyPendingLoader
    AssertEqual "", GetSetting("ThinQQ", "Updater", "PendingLoaderPath", "")
    AssertTrue Len(Dir(target)) > 0 And FileLen(target) > 0, "alvo intacto"
    PendingRestore
    Exit Sub
Fail:
    PendingFail
End Sub

Public Sub Test_Core_ApplyPendingLoader_NoLoaderRegistered_IsNoOpAndNeverRaises()
    Dim d As String, pend As String
    PendingSetup
    On Error GoTo Fail
    d = TempDir()
    pend = d & "\p.xlam"
    WriteText pend, "x"
    Core_RegisterLoader "", ""
    SaveSetting "ThinQQ", "Updater", "PendingLoaderPath", pend
    Core_ApplyPendingLoader
    AssertTrue Len(Dir(pend)) > 0, "sem loader registrado nada é feito"
    AssertEqual pend, GetSetting("ThinQQ", "Updater", "PendingLoaderPath", "")
    ' alvo em pasta inexistente + SHA valido: o erro e contido, nunca propaga
    SaveSetting "ThinQQ", "Updater", "PendingLoaderSha256", Util_Sha256File(pend)
    Core_RegisterLoader "", d & "\sem-pasta\ThinQQ.xlam"
    Core_ApplyPendingLoader
    PendingRestore
    Exit Sub
Fail:
    PendingFail
End Sub

Public Sub Test_Util_Sha256File_KnownAndMissing()
    Dim d As String
    d = TempDir()
    Dim f As Integer
    f = FreeFile: Open d & "\a.txt" For Binary As #f: Put #f, , CByte(97): Put #f, , CByte(98): Put #f, , CByte(99): Close #f
    AssertEqual "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", Util_Sha256File(d & "\a.txt")
    AssertEqual "", Util_Sha256File(d & "\nao-existe.txt")
End Sub

Private Function ReadText(path As String) As String
    Dim f As Integer, ln As String
    f = FreeFile
    Open path For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        ReadText = ReadText & ln
    Loop
    Close #f
End Function

Public Sub Test_Core_ApplyPendingLoader_BusyGuard_SkipsReentrantCall()
    Dim d As String, pend As String, target As String
    PendingSetup
    On Error GoTo Fail
    d = TempDir()
    pend = d & "\pend.xlam": target = d & "\ThinQQ.xlam"
    MakeAddinFile pend
    WriteText target, "intacto"
    Core_RegisterLoader "", target
    SaveSetting "ThinQQ", "Updater", "PendingLoaderPath", pend
    SaveSetting "ThinQQ", "Updater", "PendingLoaderSha256", Util_Sha256File(pend)
    Core_TestSetApplyBusy True
    Core_ApplyPendingLoader
    AssertTrue Len(Dir(pend)) > 0, "chamada reentrante não faz nada"
    AssertEqual pend, GetSetting("ThinQQ", "Updater", "PendingLoaderPath", "")
    AssertEqual "intacto", ReadText(target)
    Core_TestSetApplyBusy False
    Core_ApplyPendingLoader
    AssertEqual 0, Len(Dir(pend)), "sem o flag, a troca acontece"
    PendingRestore
    Exit Sub
Fail:
    PendingFail
End Sub

' ---- C1: vida do bloco por referencia (copias e graficos movidos para planilha de grafico) ----

Private Function SeriesSum(s As Series) As Double
    Dim v As Variant, i As Long
    v = s.Values
    For i = LBound(v) To UBound(v)
        SeriesSum = SeriesSum + Util_ToDouble(v(i))
    Next
End Function

Public Sub Test_Core_Cleanup_DuplicatedChart_KeepsBlockAfterOriginalDeleted()
    Dim wb As Workbook, ws As Worksheet, dup As Object, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:B5").Select   ' uma coluna de valores: cascata de uma serie (varias colunas = varias series, Tarefa 4)
    Core_Insert "waterfall"
    Set dup = ws.ChartObjects(1).Duplicate
    ws.ChartObjects(2).Name = "Gráfico 2"       ' como o Excel nomeia uma cópia colada (o Duplicate por código mantém o nome)
    AssertEqual 2, ws.ChartObjects.Count
    ws.ChartObjects(1).Delete          ' o original (nome tqq_<id>) some; a copia continua apontando para o bloco
    AssertEqual 1, ws.ChartObjects.Count
    AssertEqual 1, TqqNameCount(wb), "bloco ainda existe antes da limpeza"
    Core_Refresh
    AssertEqual 1, TqqNameCount(wb), "o bloco da copia nao pode ser apagado"
    Set cht = ws.ChartObjects(1).Chart
    AssertTrue SeriesSum(Chart_FindSeries(cht, "tqq_uppos")) > 0, "a copia continua com dados"
    ' inserir outro grafico tambem aciona a limpeza
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    AssertEqual 2, TqqNameCount(wb), "copia + novo grafico"
    AssertTrue SeriesSum(Chart_FindSeries(cht, "tqq_uppos")) > 0, "a copia continua com dados apos novo Insert"
    Teardown wb
End Sub

Public Sub Test_Core_Cleanup_ChartMovedToChartSheet_KeepsBlock()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, ch2 As Chart
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:B5").Select   ' idem
    Core_Insert "waterfall"
    Set cht = ws.ChartObjects(1).Chart
    cht.Location Where:=xlLocationAsNewSheet, Name:="GrafSheet"
    AssertEqual 0, ws.ChartObjects.Count
    AssertEqual 1, wb.Charts.Count
    Core_Refresh
    AssertEqual 1, TqqNameCount(wb), "bloco do grafico movido preservado"
    Set ch2 = wb.Charts(1)
    AssertTrue SeriesSum(Chart_FindSeries(ch2, "tqq_uppos")) > 0, "dados preservados"
    ws.Activate
    ws.Range("A1:C5").Select
    Core_Insert "clustered"
    AssertEqual 1, TqqNameCount(wb) - 1, "novo bloco + bloco movido"
    AssertTrue SeriesSum(Chart_FindSeries(wb.Charts(1), "tqq_uppos")) > 0, "dados preservados apos Insert"
    Teardown wb
End Sub

Public Sub Test_Core_Cleanup_DeletedChart_StillRemovedWhenOtherChartsExist()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:B5").Select   ' idem
    Core_Insert "waterfall"
    ws.Range("A1:B5").Select   ' idem
    Core_Insert "waterfall"
    AssertEqual 2, TqqNameCount(wb)
    ws.ChartObjects(1).Delete
    Core_Refresh
    AssertEqual 1, TqqNameCount(wb), "so o bloco do grafico excluido sai"
    AssertTrue SeriesSum(Chart_FindSeries(ws.ChartObjects(1).Chart, "tqq_uppos")) > 0
    Teardown wb
End Sub

' ---- I2: a nova fonte tem de estar na mesma pasta de trabalho do grafico ----

Public Sub Test_Core_RebuildFromRange_RangeFromOtherWorkbook_Rejected()
    Dim wb As Workbook, wb2 As Workbook, ws As Worksheet, ws2 As Worksheet, cht As Chart, nm As String
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set wb2 = NewTempWorkbook()
    Set ws2 = wb2.Worksheets(1): ws2.Name = "Dados"          ' mesmo nome de planilha da pasta do grafico
    ws2.Range("A1:B1").Value = Array("Ano", "Outro")
    ws2.Range("A2:B2").Value = Array(2022, 999)
    ws2.Range("A3:B3").Value = Array(2023, 888)
    wb.Activate
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    nm = ws.ChartObjects(1).Name
    Core_ClearLastNotice
    AssertFalse Core_RebuildFromRange(cht, ws2.Range("A1:B3")), "intervalo de outra pasta deve ser recusado"
    AssertEqual "Selecione uma tabela na mesma pasta de trabalho do gráfico.", Core_LastNotice()
    AssertEqual 1, ws.ChartObjects.Count
    AssertEqual nm, ws.ChartObjects(1).Name, "o grafico original continua intacto"
    AssertEqual 1, TqqNameCount(wb)
    AssertEqual 0, TqqNameCount(wb2), "nada criado na outra pasta"
    wb2.Close SaveChanges:=False
    Teardown wb
End Sub

' ---- I3: linhas inseridas/removidas na origem dessincronizam o bloco: detectar e avisar uma unica vez ----

Private Sub CheckMismatch(kind As String)
    Dim wb As Workbook, ws As Worksheet, cht As Chart, msg As String
    msg = "A tabela de origem mudou de tamanho. Use 'Alterar dados' para reconstruir o gráfico (ou reinsira o gráfico)."
    Core_OnUnload
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert kind
    Set cht = Core_ActiveChart()
    AssertFalse Core_ChartSourceMismatch(cht), kind & ": sem mudanca nao ha divergencia"
    Core_ClearLastNotice
    Core_SafeRefresh cht
    AssertEqual 0, Core_NoticeCount(), kind & ": sem aviso quando em sincronia"
    ws.Rows(3).Insert                           ' (com observadores ativos o evento de calculo ja pode avisar aqui)
    ws.Range("A3:C3").Value = Array(2022.5, 105, 55)
    AssertTrue Core_ChartSourceMismatch(cht), kind & ": linha inserida deve ser detectada"
    Core_SafeRefresh cht
    AssertEqual msg, Core_LastNotice(), kind & ": aviso apos a mudanca"
    Core_SafeRefresh cht
    Core_Refresh
    AssertEqual 1, Core_NoticeCount(), kind & ": avisa exatamente uma vez"
    Core_OnUnload                               ' nova sessao: avisa de novo
    Core_ClearLastNotice
    Core_SafeRefresh cht
    AssertEqual 1, Core_NoticeCount(), kind & ": reiniciado em Core_OnUnload"
    ws.Rows(3).Delete
    ws.Rows(3).Delete                           ' remove outra linha: agora a origem tem menos linhas que o bloco
    AssertTrue Core_ChartSourceMismatch(cht), kind & ": linha removida tambem diverge"
    Teardown wb
End Sub

Public Sub Test_Core_SourceMismatch_Waterfall_RowInserted_NoticedOnce()
    CheckMismatch "waterfall"
End Sub

Public Sub Test_Core_SourceMismatch_Stacked_RowInserted_NoticedOnce()
    CheckMismatch "stacked"
End Sub

' ---- I4: graficos 100% rotulam os segmentos com a participacao ----

Private Function Fixture100(wb As Workbook) As Chart
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Ano", "Sul", "Sudeste", "Nordeste")
    ws.Range("A2:D2").Value = Array(2022, 10, 20, 5)
    ws.Range("A3:D3").Value = Array(2023, "n/d", 30, 10)
    ws.Range("A4:D4").Value = Array(2024, 0, 0, 0)
    ws.Range("A1:D4").Select
    Core_Insert "stacked100"
    Set Fixture100 = Core_ActiveChart()
End Function

Public Sub Test_Core_Stacked100_SegmentLabels_ShowShare()
    Dim wb As Workbook, cht As Chart, tot As Series, d As String, body As Range
    d = Application.International(xlDecimalSeparator)
    Set wb = NewTempWorkbook()
    Set cht = Fixture100(wb)
    AssertEqual "29%", cht.SeriesCollection(1).Points(1).DataLabel.Text, "10/35"
    AssertEqual "57%", cht.SeriesCollection(2).Points(1).DataLabel.Text, "20/35"
    AssertEqual "14%", cht.SeriesCollection(3).Points(1).DataLabel.Text, "5/35"
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertEqual "35", tot.Points(1).DataLabel.Text, "o total continua absoluto"
    AssertEqual xlLabelPositionCenter, cht.SeriesCollection(1).Points(1).DataLabel.Position
    ' texto vira 0 (rotulo vazio pelo formato) sem erros; participacao das demais usa o total sem o texto (40)
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual 0, body.Cells(3, 4).Value, "texto = participacao 0"
    AssertEqual "", cht.SeriesCollection(1).Points(2).DataLabel.Text, "zero oculto"
    AssertEqual "75%", cht.SeriesCollection(2).Points(2).DataLabel.Text, "30/40"
    ' total 0: sem #DIV/0!
    AssertEqual 0, body.Cells(4, 4).Value
    AssertFalse IsError(body.Cells(4, 5).Value), "linha de total zero nao gera #DIV/0!"
    Teardown wb
End Sub

Public Sub Test_Core_Stacked100_NumFmt_AppliesToTotals_Pct1ToShares()
    Dim wb As Workbook, cht As Chart, tot As Series, d As String
    d = Application.International(xlDecimalSeparator)
    Set wb = NewTempWorkbook()
    Set cht = Fixture100(wb)
    Set tot = Chart_FindSeries(cht, "tqq_total")
    Core_SetOption "numfmt", "dec1"
    AssertEqual "35" & d & "0", tot.Points(1).DataLabel.Text, "dec1 vale para os totais"
    AssertEqual "29%", cht.SeriesCollection(1).Points(1).DataLabel.Text, "participacao continua pct0"
    Core_SetOption "numfmt", "pct1"
    AssertEqual "28" & d & "6%", cht.SeriesCollection(1).Points(1).DataLabel.Text, "pct1 = participacao com uma casa"
    AssertEqual "35", tot.Points(1).DataLabel.Text, "totais continuam absolutos (inteiro) com pct1"
    Core_SetOption "numfmt", "int"
    AssertEqual "29%", cht.SeriesCollection(1).Points(1).DataLabel.Text, "volta a pct0"
    Teardown wb
End Sub

' ---- M12: Watchers_Rehook redesenha os graficos recem-observados ----

Public Sub Test_Watchers_Rehook_RefreshesAnnotations()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Watchers_Stop
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    Set cht = ws.ChartObjects(1).Chart
    Core_AddAnnotation "cagr", 1, 4
    Watchers_Stop                                   ' sem observadores: mudancas nao redesenham
    ws.Range("B5").Value = 250                      ' total 2025 = 340
    AssertEqual "CAGR +17,0%", cht.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text, "rotulo ainda desatualizado"
    Watchers_Rehook                                 ' ao observar de novo, atualiza
    AssertEqual "CAGR +31,4%", cht.Shapes("tqq_ann_cagr_1_lbl").TextFrame2.TextRange.Text
    Teardown wb
End Sub

' ---- Tarefa 0: limpeza de blocos conta rotulos ligados ----

' Tarefa 0: um 100% sem totais so referencia o bloco pelos rotulos ligados (DataLabel.Formula)
Public Sub Test_Core_Cleanup_Stacked100NoTotals_CopyKeepsBlock()
    Dim wb As Workbook, ws As Worksheet, dup As Object, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked100"
    Core_SetToggle "totals", False
    AssertTrue Chart_FindSeries(Core_ActiveChart(), "tqq_total") Is Nothing
    Set dup = ws.ChartObjects(1).Duplicate
    ws.ChartObjects(2).Name = "Gráfico 2"       ' como o Excel nomeia uma copia colada
    ws.ChartObjects(1).Delete
    AssertEqual 1, TqqNameCount(wb)
    Core_Refresh
    AssertEqual 1, TqqNameCount(wb), "bloco referenciado so por rotulos ligados nao pode ser apagado"
    Set cht = ws.ChartObjects(1).Chart
    AssertEqual Format$(100 / 150, "0%"), cht.SeriesCollection(1).Points(1).DataLabel.Text, "a copia continua mostrando a participacao (100/150)"
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    AssertEqual 2, TqqNameCount(wb), "copia + novo grafico"
    AssertEqual Format$(100 / 150, "0%"), cht.SeriesCollection(1).Points(1).DataLabel.Text, "o novo bloco nao reaproveitou as celulas da copia"
    Teardown wb
End Sub

' Grupo de formas: o grafico dentro do grupo continua referenciando o bloco (codigo da Fase 1 sem teste dedicado)
Public Sub Test_Core_Cleanup_ChartInsideGroup_KeepsBlock()
    Dim wb As Workbook, ws As Worksheet, dup As Object, rect As Shape
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "waterfall"
    Set dup = ws.ChartObjects(1).Duplicate
    ws.ChartObjects(2).Name = "Gráfico 2"
    ws.ChartObjects(1).Delete
    Set rect = ws.Shapes.AddShape(msoShapeRectangle, 10, 400, 50, 20)
    ws.Shapes.Range(Array("Gráfico 2", rect.Name)).Group
    Core_Refresh
    AssertEqual 1, TqqNameCount(wb), "bloco de grafico dentro de grupo preservado"
    Teardown wb
End Sub

' Falha fechada da limpeza nao impede o Insert: a limpeza aborta (orfao intacto, nada apagado) e o novo bloco fica
' abaixo do orfao, sem reaproveitar as celulas dele. Depois, sem falha, a limpeza remove so o orfao.
Public Sub Test_Core_Insert_WhileCleanupAborts_StillInsertsAndDoesNotOverlapOrphan()
    Dim wb As Workbook, ws As Worksheet, orphan As Range, nb As Range, cht As Chart
    Dim en As Long, es As String, ed As String
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.ChartObjects.Add 400, 10, 200, 150                       ' um grafico qualquer: ha leitura a fazer
    Set orphan = CalcSheet_AllocateBlock(wb, "orphan", "stacked", "x", 4, 3)
    orphan.Cells(2, 1).Value = "marca"
    ws.Range("A1:C5").Select
    On Error GoTo Cleanup
    Util_FailChartReadSiteForTest = READFAIL_SHAPE
    Core_Insert "stacked"
    Util_FailChartReadSiteForTest = 0
    AssertTrue CalcSheet_CleanupReadFailed(), "a limpeza do Insert abortou"
    AssertEqual 2, ws.ChartObjects.Count, "o grafico foi inserido"
    AssertEqual 2, TqqNameCount(wb), "orfao mantido + bloco novo"
    Set cht = Core_ActiveChart()
    AssertTrue Not cht Is Nothing
    Set nb = Registry_Block(cht)
    AssertTrue Application.Intersect(nb, orphan) Is Nothing, "o novo bloco nao reaproveita o orfao"
    AssertTrue nb.Row > orphan.Row + orphan.Rows.Count - 1, "o novo bloco fica abaixo do orfao"
    AssertEqual "marca", orphan.Cells(2, 1).Value, "o orfao nao foi sobrescrito"
    Core_Refresh                                                ' sem falha: remove so o orfao
    AssertFalse CalcSheet_CleanupReadFailed()
    AssertEqual 1, TqqNameCount(wb), "so o bloco do novo grafico"
    AssertTrue CalcSheet_FindBlock(wb, "orphan") Is Nothing
    AssertTrue Not Registry_Block(cht) Is Nothing
    Teardown wb
    Exit Sub
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailChartReadSiteForTest = 0
    Teardown wb
    Err.Raise en, es, ed
End Sub

Public Sub Test_Core_Supports_FollowsActiveChart()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    AssertTrue Core_Supports("totals"): AssertTrue Core_Supports("changesource")
    AssertFalse Core_Supports("secondary"): AssertFalse Core_Supports("pieout")
    AssertEqual "4", Core_GetOption("n"), "numero de categorias persistido"
    ws.Range("A1").Select
    AssertFalse Core_Supports("totals"): AssertFalse Core_Supports("changesource")
    Teardown wb
End Sub

' Build Test: o log vai para %TEMP%\ThinQQ-test-logs, nunca para o log real do usuario.
Public Sub Test_Errors_Log_TestBuildWritesToTempFolder()
    Dim marker As String, p As String, f As Integer, txt As String
    AssertEqual Environ$("TEMP") & "\ThinQQ-test-logs", Errors_LogDir()
    Randomize
    marker = "tqq-log-probe-" & Format$(Now, "hhnnss") & "-" & CStr(Int(Rnd * 100000))
    Errors_Log marker
    p = Errors_LogDir() & "\thinqq.log"
    AssertTrue Len(Dir(p)) > 0, "log de teste criado em " & p
    f = FreeFile
    Open p For Binary Access Read As #f
    txt = Space$(LOF(f))
    Get #f, , txt
    Close #f
    AssertTrue InStr(txt, marker) > 0, "a linha foi para o log de teste"
End Sub

' Permanente: tipo que nao existe e recusado com o aviso de tipo desconhecido (nada e criado), inclusive uma lista de
' tipos separada por virgula (a matriz nao aceita uma lista como se fosse um tipo).
Public Sub Test_Core_Insert_UnknownKind_NotifiesAndCreatesNothing()
    Dim wb As Workbook, ws As Worksheet, kind As Variant
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    For Each kind In Array("xyz", "stacked,clustered", "barstacked,")
        Core_ClearLastNotice
        Core_Insert CStr(kind)
        AssertEqual 0, ws.ChartObjects.Count, kind & ": nenhum grafico criado"
        AssertEqual 0, TqqNameCount(wb), kind & ": nenhum bloco criado"
        AssertTrue InStr(Core_LastNotice(), "Tipo de gráfico desconhecido: " & kind) > 0, kind & ": " & Core_LastNotice()
    Next
    Teardown wb
End Sub

' Barras pela API: tipo, totais (liga/desliga/religa), formato e espacamento valem para o grafico horizontal.
Public Sub Test_Core_Insert_Bars_ThroughApi()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "barstacked"
    Set cht = Core_ActiveChart()
    AssertEqual "barstacked", Registry_ChartType(cht)
    AssertTrue Core_GetToggle("totals")
    Core_SetToggle "totals", False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    Core_SetToggle "totals", True
    AssertTrue Not Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertEqual 2, cht.Legend.LegendEntries.Count, "totais religados continuam fora da legenda"
    Core_SetOption "numfmt", "dec1"
    AssertEqual "150" & Application.International(xlDecimalSeparator) & "0", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    Core_SetOption "gap", "120"
    AssertEqual 120, cht.ChartGroups(1).GapWidth
    Teardown wb
End Sub

' Barras agrupadas e 100% pela API e reconstruidas por Core_RebuildFromRange: mesmo tipo, totais XY e opcoes.
Public Sub Test_Core_RebuildFromRange_Bars_KeepsKindAndTotalsXY()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, nc As Chart, kind As Variant
    For Each kind In Array("barstacked", "barstacked100", "barclustered")
        Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
        ws.Range("E1:G1").Value = Array("Ano", "Sul", "Norte")
        ws.Range("E2:G2").Value = Array(2022, 10, 5): ws.Range("E3:G3").Value = Array(2023, 20, 6)
        ws.Range("A1:C5").Select
        Core_Insert CStr(kind)
        Set cht = Core_ActiveChart()
        Core_SetOption "numfmt", "dec1"
        AssertTrue Core_RebuildFromRange(cht, ws.Range("E1:G3")), kind & ": reconstruiu"
        Set nc = Core_ActiveChart()
        AssertEqual CStr(kind), Registry_ChartType(nc)
        AssertEqual "2", Registry_GetOption(nc, "n")
        AssertFalse Core_ChartSourceMismatch(nc), kind & ": n acompanha a nova tabela"
        If kind = "barclustered" Then
            AssertTrue Chart_FindSeries(nc, "tqq_total") Is Nothing
        Else
            AssertTrue Not Chart_FindSeries(nc, "tqq_total") Is Nothing, kind & ": totais refeitos"
            AssertEqual xlSecondary, Chart_FindSeries(nc, "tqq_total").AxisGroup
            AssertEqual 2, nc.Legend.LegendEntries.Count, kind & ": legenda sem a auxiliar"
        End If
        Teardown wb
    Next
End Sub

' A opcao n (numero de categorias) vale para a cascata tambem; sem ela (grafico de versao antiga) a divergencia
' continua sendo medida pelo tamanho do bloco.
Public Sub Test_Core_Insert_Waterfall_PersistsN_And_LegacyChartWithoutN()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Início", 100)
    ws.Range("A3:B3").Value = Array("A", -30)
    ws.Range("A4:B4").Value = Array("Fim", "e")
    ws.Range("A1").Select
    Core_Insert "waterfall"
    Set cht = Core_ActiveChart()
    AssertEqual "3", Core_GetOption("n")
    AssertFalse Core_ChartSourceMismatch(cht)
    Registry_Block(cht).Cells(1, 4).Value = "numfmt=int"          ' grafico antigo: sem a opcao n
    AssertEqual "", Core_GetOption("n")
    AssertFalse Core_ChartSourceMismatch(cht), "sem n: compara com o tamanho do bloco"
    ws.Rows(3).Insert
    AssertTrue Core_ChartSourceMismatch(cht), "sem n: linha inserida ainda e detectada"
    Teardown wb
End Sub

' A opcao de legenda: o "fim da linha" so existe em linha e combinacao e a recusa (7002) nao persiste a opcao; "none"
' desliga a legenda; o toggle "secondary" so vale na combinacao.
Public Sub Test_Core_Legend_And_Secondary_Options()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    Core_ClearLastNotice
    Core_SetOption "legend", "lineend"
    AssertEqual "Rótulos no fim da linha só existem em gráficos de linha e combinação.", Core_LastNotice()
    AssertEqual "right", Core_GetOption("legend"), "nao persiste"
    Core_SetOption "legend", "none"
    AssertFalse Core_ActiveChart().HasLegend
    Core_SetToggle "secondary", True
    AssertFalse Core_GetToggle("secondary"), "secondary so em combinacao"
    Teardown wb
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Tarefa 11: API unificada para todos os tipos (toggles e opcoes pela matriz, reconstrucao generica, setas e colchete)
' ---------------------------------------------------------------------------------------------------------------------

' Tabela de dados adequada a cada tipo, escrita a partir de A1 de ws; devolve a faixa.
Private Function FixtureFor(ws As Worksheet, kind As String) As Range
    Select Case Kinds_Family(kind)
        Case "waterfall"
            ws.Range("A1:C1").Value = Array("Etapa", "Loja", "Online")
            ws.Range("A2:C2").Value = Array("Inicio", 100, 50): ws.Range("A3:C3").Value = Array("Ajuste", -30, 10)
            ws.Range("A4:C4").Value = Array("Fim", "e", "e")
            Set FixtureFor = ws.Range("A1:C4")
        Case "mekko"
            ws.Range("A1:D1").Value = Array("Regiao", "Largura", "Eletro", "Moveis")
            ws.Range("A2:D2").Value = Array("Sul", 50, 40, 10): ws.Range("A3:D3").Value = Array("Norte", 30, 20, 10)
            If kind = "mekko" Then ws.Columns("B").Delete          ' percentual: sem coluna de largura
            Set FixtureFor = ws.Range("A1").CurrentRegion
        Case "gantt"
            ws.Range("A1:C1").Value = Array("Atividade", "Início", "Fim")
            ws.Range("A2").Value = "A": ws.Range("B2").Value = DateSerial(2026, 3, 2): ws.Range("C2").Value = DateSerial(2026, 3, 13)
            ws.Range("A3").Value = "B": ws.Range("B3").Value = DateSerial(2026, 3, 16)
            Set FixtureFor = ws.Range("A1:C3")
        Case "pie"
            ws.Range("A1:B1").Value = Array("Produto", "Vendas")
            ws.Range("A2:B2").Value = Array("Eletro", 50): ws.Range("A3:B3").Value = Array("Moveis", 30)
            Set FixtureFor = ws.Range("A1:B3")
        Case Else
            ws.Range("A1:D1").Value = Array("Ano", "Sul", "Sudeste", "Nordeste")
            ws.Range("A2:D2").Value = Array(2022, 10, 20, 5): ws.Range("A3:D3").Value = Array(2023, 12, 22, 6)
            ws.Range("A4:D4").Value = Array(2024, 15, 25, 8)
            Set FixtureFor = ws.Range("A1:D4")
    End Select
End Function

' A mesma tabela do tipo com UMA categoria so (cabecalho + uma linha).
Private Function OneRowFixtureFor(ws As Worksheet, kind As String) As Range
    Dim rng As Range
    Set rng = FixtureFor(ws, kind)
    ws.Rows("3:" & rng.Rows.Count).Delete
    Set OneRowFixtureFor = ws.Range("A1").CurrentRegion
End Function

' O que o registro guarda das escolhas do usuario (alternadores e opcoes), em uma linha: a comparacao antes e depois de "Alterar dados"
' mostra se alguma escolha foi sobrescrita pelo estado real do grafico novo (ex.: indicador 100% desligado voltando ligado).
Private Function ChoicesSnapshot() As String
    Dim k As Variant
    For Each k In Array("totals", "labels", "connectors", "ind100", "vline", "secondary", "pieout", "numfmt", "palette", "gap", "legend", "n")
        ChoicesSnapshot = ChoicesSnapshot & k & "=" & Core_GetOption(CStr(k)) & ";"
    Next
End Function

' Formas de anotacao (setas, colchetes, indicador) do grafico.
Private Function AnnotationShapeCount(c As Chart) As Long
    Dim i As Long
    For i = 1 To c.Shapes.Count
        If Left$(c.Shapes(i).Name, Len(ANN_PREFIX)) = ANN_PREFIX Then AnnotationShapeCount = AnnotationShapeCount + 1
    Next
End Function

' Leituras defensivas (uma propriedade que o Excel recusa para o tipo de grupo/serie vale False; nunca dentro de uma condicao de If).
Private Function GroupHasSeriesLines(g As ChartGroup) As Boolean
    On Error Resume Next
    GroupHasSeriesLines = g.HasSeriesLines
End Function

Private Function SeriesIsSecondary(s As Series) As Boolean
    On Error Resume Next
    SeriesIsSecondary = (s.AxisGroup = xlSecondary)
End Function

Private Function SeriesHasLabels(s As Series) As Boolean
    On Error Resume Next
    SeriesHasLabels = s.HasDataLabels
End Function

' O ESTADO REAL do grafico (lido do Excel: series, grupos e formas; nunca as opcoes gravadas) depois de o teste ligar todo recurso que o tipo
' suporta e desligar Rotulos e Indicador 100%: totais (serie tqq_total), rotulos (nenhuma serie de dados com rotulo e nenhuma auxiliar
' tqq_lbl*), conectores (linhas de serie dos grupos empilhados; na cascata, as formas tqq_ann_conn_*), eixo secundario (a linha do combo),
' linha de valor (tqq_vline) e indicador 100% (formas). Vale antes E depois de "Alterar dados".
Private Sub AssertRealState(c As Chart, kind As String, tag As String)
    Dim s As Series, g As ChartGroup, i As Long, labelsOn As Boolean, sec As Boolean, conn As Boolean
    For Each s In c.SeriesCollection
        If Style_IsHelperSeries(s) Then
            If Left$(s.Name, 7) = "tqq_lbl" Then labelsOn = True
        Else
            If SeriesHasLabels(s) Then labelsOn = True
            If SeriesIsSecondary(s) Then sec = True
        End If
    Next
    If Kinds_Family(kind) = "waterfall" Then
        For i = 1 To c.Shapes.Count
            If Left$(c.Shapes(i).Name, Len(ANN_PREFIX & "conn_")) = ANN_PREFIX & "conn_" Then conn = True
        Next
    Else
        For Each g In c.ChartGroups
            If GroupHasSeriesLines(g) Then conn = True
        Next
    End If
    ' (na cascata tqq_total sao as barras de total do proprio grafico, nao o alternador Totais)
    If Kinds_Family(kind) <> "waterfall" Then AssertEqual Kinds_Supports(kind, "totals"), Not Chart_FindSeries(c, "tqq_total") Is Nothing, tag & ": totais (serie tqq_total)"
    AssertFalse labelsOn, tag & ": rotulos desligados pelo usuario continuam desligados (serie com rotulo ou tqq_lbl*)"
    AssertEqual Kinds_Supports(kind, "connectors"), conn, tag & ": conectores"
    AssertEqual Kinds_Supports(kind, "secondary"), sec, tag & ": eixo secundario"
    AssertEqual Kinds_Supports(kind, "vline"), Not Chart_FindSeries(c, "tqq_vline") Is Nothing, tag & ": linha de valor"
    AssertFalse Ann_HasInd100(c), tag & ": indicador 100% desligado pelo usuario continua desligado"
End Sub

' Todo tipo: insere pela API, exercita cada toggle/opcao conforme a matriz (suportado = funciona; nao suportado = aviso,
' nada persiste), reconstroi mantendo o tipo e as opcoes, e aceita/recusa setas pela matriz.
Public Sub Test_Core_AllKinds_InsertTogglesOptionsRebuild()
    Dim kinds As Variant, toggles As Variant, k As Long, t As Long, wb As Workbook, ws As Worksheet, rng As Range, c As Chart, kind As String, key As String
    Dim before As String, ch As Variant
    kinds = Array("stacked", "clustered", "stacked100", "line", "waterfall", "barstacked", "barclustered", "barstacked100", "barwaterfall", _
                  "area", "area100", "combo", "mekko", "mekkounits", "gantt", "pie", "doughnut")
    toggles = Array("totals", "labels", "connectors", "ind100", "vline", "secondary", "pieout")
    For k = LBound(kinds) To UBound(kinds)
        kind = kinds(k)
        Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
        Set rng = FixtureFor(ws, kind)
        rng.Select
        Core_ClearLastNotice
        Core_Insert kind
        Set c = Core_ActiveChart()
        AssertTrue Not c Is Nothing, kind & ": nao inseriu (" & Core_LastNotice() & ")"
        AssertEqual kind, Registry_ChartType(c)
        For t = LBound(toggles) To UBound(toggles)
            key = toggles(t)
            Core_ClearLastNotice
            If Kinds_Supports(kind, key) Then
                Core_SetToggle key, True
                AssertTrue Core_GetToggle(key), kind & "/" & key & " ligado"
                Core_SetToggle key, False
                AssertFalse Core_GetToggle(key), kind & "/" & key & " desligado"
                Core_SetToggle key, True
            Else
                Core_SetToggle key, True
                AssertEqual "'" & Kinds_FeatureLabel(key) & "' não se aplica a este tipo de gráfico.", Core_LastNotice(), kind & "/" & key
                AssertFalse Core_GetToggle(key), kind & "/" & key & " nao persiste"
            End If
        Next
        Core_ClearLastNotice
        Core_SetOption "numfmt", "dec1"
        If Kinds_Supports(kind, "numfmt") Then
            AssertEqual "dec1", Core_GetOption("numfmt"), kind
        Else
            AssertEqual "'Formato' não se aplica a este tipo de gráfico.", Core_LastNotice(), kind & "/numfmt"
            AssertTrue Core_GetOption("numfmt") <> "dec1", kind & "/numfmt nao persiste"
        End If
        Core_SetOption "palette", "cinza": AssertEqual "cinza", Core_GetOption("palette"), kind
        Core_ClearLastNotice
        Core_SetOption "gap", "80"
        If Kinds_Supports(kind, "gap") Then
            AssertEqual "80", Core_GetOption("gap"), kind & "/gap"
            AssertEqual 80, c.ChartGroups(1).GapWidth, kind & "/gap aplicado"
        Else
            AssertEqual "Espaço entre colunas não se aplica a este tipo de gráfico.", Core_LastNotice(), kind & "/gap"
            AssertTrue Core_GetOption("gap") <> "80", kind & "/gap nao persiste"
        End If
        If Kinds_SupportsChart(c, "legend") Then Core_SetOption "legend", "none": AssertFalse c.HasLegend, kind & "/legend"      ' (a cascata: so com varias series)
        ' escolhas do usuario DIFERENTES do padrao do Insert (rotulos e indicador 100% desligados): "Alterar dados" as mantem
        For Each ch In Array("labels", "ind100")
            If Kinds_Supports(kind, CStr(ch)) Then Core_SetToggle CStr(ch), False
        Next
        before = ChoicesSnapshot()
        AssertRealState c, kind, kind & ": antes de Alterar dados"
        Core_ClearLastNotice
        AssertTrue Core_RebuildFromRange(c, rng), kind & ": rebuild (" & Core_LastNotice() & ")"
        Set c = Core_ActiveChart()
        AssertEqual kind, Registry_ChartType(c), kind & ": tipo apos rebuild"
        AssertEqual before, ChoicesSnapshot(), kind & ": opcoes persistidas apos rebuild"
        AssertRealState c, kind, kind & ": apos Alterar dados"          ' o estado REAL do grafico novo, nao so as opcoes copiadas
        If Kinds_Supports(kind, "numfmt") Then AssertEqual "dec1", Core_GetOption("numfmt"), kind & ": numfmt apos rebuild"
        AssertEqual "cinza", Core_GetOption("palette"), kind & ": paleta apos rebuild"
        If Kinds_Supports(kind, "vline") Then AssertTrue Not Chart_FindSeries(c, "tqq_vline") Is Nothing, kind & ": linha de valor apos rebuild"
        If Kinds_Supports(kind, "ind100") Then AssertFalse Ann_HasInd100(c), kind & ": indicador desligado pelo usuario continua desligado"
        If Kinds_SupportsChart(c, "legend") Then AssertFalse c.HasLegend, kind & ": legenda 'nenhuma' apos rebuild"
        AssertEqual 1, TqqNameCount(wb), kind & ": um bloco"
        Core_ClearLastNotice
        Core_AddAnnotation "cagr", 1, 2
        If Kinds_Supports(kind, "cagr") Then
            AssertTrue Not c.Shapes("tqq_ann_cagr_1") Is Nothing, kind & ": CAGR"
        Else
            AssertEqual "A seta CAGR não está disponível para este tipo de gráfico.", Core_LastNotice(), kind & ": recusa CAGR"
        End If
        Teardown wb
    Next
End Sub

Public Sub Test_Core_AddAnnotation_Unsupported_NotifiesBeforeForm()
    Dim wb As Workbook, ws As Worksheet, rng As Range
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Set rng = FixtureFor(ws, "stacked")
    rng.Select
    Core_Insert "stacked"
    Core_ClearLastNotice
    Core_AddAnnotation "bracket"                       ' sem indices: nao pode abrir formulario nem falhar
    AssertEqual "O colchete só está disponível em gráficos Gantt.", Core_LastNotice()
    AssertEqual 1, Core_NoticeCount(), "um unico aviso (conta desde Core_ClearLastNotice)"
    Teardown wb
End Sub

' A recusa da matriz vem ANTES de qualquer outra verificacao (formulario de categorias e "minimo de categorias"): uma pizza ou um
' Gantt de UMA linha, o Mekko e os 100% recebem a mensagem do tipo, sem indices e sem formulario; os tipos que aceitam seguem as regras
' de indices ("pelo menos 2 categorias" so para quem suporta a seta).
Public Sub Test_Core_AddAnnotation_RefusalComesBeforeCountAndForm()
    Dim wb As Workbook, ws As Worksheet, cases As Variant, i As Long, kind As String, ann As String, want As String, c As Chart, one As Boolean, shapesBefore As Long
    ' tipo; uma categoria so?; anotacao; aviso esperado
    cases = Array( _
        Array("pie", True, "cagr", "A seta CAGR não está disponível para este tipo de gráfico."), _
        Array("pie", True, "diff", "A seta de diferença não está disponível para este tipo de gráfico."), _
        Array("pie", True, "bracket", "O colchete só está disponível em gráficos Gantt."), _
        Array("doughnut", False, "cagr", "A seta CAGR não está disponível para este tipo de gráfico."), _
        Array("mekko", False, "cagr", "A seta CAGR não está disponível para este tipo de gráfico."), _
        Array("mekko", False, "diff", "A seta de diferença não está disponível para este tipo de gráfico."), _
        Array("mekkounits", False, "diff", "A seta de diferença não está disponível para este tipo de gráfico."), _
        Array("stacked100", False, "diff", "A seta de diferença não está disponível para gráficos 100%."), _
        Array("gantt", True, "cagr", "A seta CAGR não está disponível para este tipo de gráfico."), _
        Array("gantt", True, "diff", "A seta de diferença não está disponível para este tipo de gráfico."), _
        Array("clustered", False, "bracket", "O colchete só está disponível em gráficos Gantt."), _
        Array("stacked", True, "cagr", "O gráfico precisa de pelo menos 2 categorias."), _
        Array("stacked", True, "diff", "O gráfico precisa de pelo menos 2 categorias."))
    For i = LBound(cases) To UBound(cases)
        kind = cases(i)(0): one = cases(i)(1): ann = cases(i)(2): want = cases(i)(3)
        Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
        If one Then OneRowFixtureFor(ws, kind).Select Else FixtureFor(ws, kind).Select
        Core_Insert kind
        Set c = Core_ActiveChart()
        AssertTrue Not c Is Nothing, kind & ": nao inseriu (" & Core_LastNotice() & ")"
        If one Then AssertEqual 1, Geom_ChartCategoryCount(c), kind & ": premissa, uma categoria"
        shapesBefore = AnnotationShapeCount(c)           ' (os 100% e o Mekko % ja trazem o indicador 100%)
        Core_ClearLastNotice
        Core_AddAnnotation ann                          ' sem indices: o formulario nunca pode ser o primeiro passo
        AssertEqual want, Core_LastNotice(), kind & "/" & ann
        AssertEqual 1, Core_NoticeCount(), kind & "/" & ann & ": um unico aviso"
        AssertEqual shapesBefore, AnnotationShapeCount(c), kind & "/" & ann & ": nenhuma forma criada"
        Teardown wb
    Next
    ' O colchete aceita a mesma atividade nas duas pontas (de = ate) e nao pede o formulario: com os indices certos cria a forma.
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    FixtureFor(ws, "gantt").Select
    Core_Insert "gantt"
    Set c = Core_ActiveChart()
    Core_ClearLastNotice
    Core_AddAnnotation "bracket", 2, 2
    AssertTrue Not c.Shapes("tqq_ann_bracket_1") Is Nothing, "colchete de uma atividade (de = ate)"
    AssertEqual 0, Core_NoticeCount(), "sem aviso"
    Teardown wb
End Sub

' Escreve a partir de K1 uma tabela de teste com "cols" colunas (categoria + valores) e 2 linhas; devolve a faixa.
Private Function BadTable(ws As Worksheet, cols As Long) As Range
    Dim j As Long
    ws.Range("K1").Value = "Etapa": ws.Range("K2").Value = "A": ws.Range("K3").Value = "B"
    For j = 2 To cols
        ws.Cells(1, 10 + j).Value = "S" & (j - 1): ws.Cells(2, 10 + j).Value = 1: ws.Cells(3, 10 + j).Value = 2
    Next
    Set BadTable = ws.Range("K1").Resize(3, cols)
End Function

' "Alterar dados" com dados que o tipo recusa (codigos 6007..6099, validacao dos dados): aviso com a mensagem do tipo, nao o erro generico;
' o grafico antigo fica intacto (mesmo tipo, mesmo bloco) e nada novo sobra.
Public Sub Test_Core_RebuildFromRange_ValidationRefusals_NoticeNotGenericError()
    Dim wb As Workbook, ws As Worksheet, cases As Variant, i As Long, kind As String, c As Chart, want As String, idBefore As String
    ' tipo; aviso esperado; colunas da tabela recusada (8 = 7 series; 2 = uma coluna de valores so)
    cases = Array( _
        Array("waterfall", "A cascata aceita no máximo 6 séries.", 8), _
        Array("barwaterfall", "A cascata aceita no máximo 6 séries.", 8), _
        Array("combo", "A combinação precisa de pelo menos 2 colunas de valores: as primeiras viram colunas empilhadas e a última vira linha.", 2), _
        Array("mekkounits", "Mekko (unidades): a segunda coluna da tabela é a largura; são necessárias pelo menos 3 colunas.", 2), _
        Array("gantt", "Gantt: a tabela precisa de 3 colunas: atividade, início e fim (datas). Deixe o fim vazio para um marco.", 2))
    For i = LBound(cases) To UBound(cases)
        kind = cases(i)(0): want = cases(i)(1)
        Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
        FixtureFor(ws, kind).Select
        Core_Insert kind
        Set c = Core_ActiveChart()
        idBefore = Registry_IdFromChart(c)
        Core_ClearLastNotice
        AssertFalse Core_RebuildFromRange(c, BadTable(ws, CLng(cases(i)(2)))), kind & ": recusa devolve False"
        AssertEqual want, Core_LastNotice(), kind & ": aviso do tipo (nao o erro generico)"
        AssertEqual 1, Core_NoticeCount(), kind & ": um unico aviso"
        AssertEqual 1, ws.ChartObjects.Count, kind & ": o grafico antigo fica"
        AssertEqual 1, TqqNameCount(wb), kind & ": nenhum bloco novo"
        Set c = Core_ActiveChart()
        AssertEqual idBefore, Registry_IdFromChart(c), kind & ": e o mesmo grafico"
        AssertEqual kind, Registry_ChartType(c)
        Teardown wb
    Next
End Sub

' numfmt, gap e legenda fora da matriz (Gantt: sem formato numerico nem legenda; linha, area, pizza e Mekko: sem espacamento; cascata de uma
' serie: sem legenda) avisam e NAO gravam (antes o Gantt gravava o formato em silencio).
Public Sub Test_Core_SetOption_UnsupportedByMatrix_NoticeNotPersisted()
    Dim wb As Workbook, ws As Worksheet, cases As Variant, i As Long, kind As String, key As String, val As String, want As String
    cases = Array( _
        Array("gantt", "numfmt", "dec1", "'Formato' não se aplica a este tipo de gráfico."), _
        Array("line", "gap", "80", "Espaço entre colunas não se aplica a este tipo de gráfico."), _
        Array("area", "gap", "80", "Espaço entre colunas não se aplica a este tipo de gráfico."), _
        Array("pie", "gap", "80", "Espaço entre colunas não se aplica a este tipo de gráfico."), _
        Array("mekko", "gap", "80", "Espaço entre colunas não se aplica a este tipo de gráfico."), _
        Array("gantt", "legend", "none", "Legenda não se aplica a este tipo de gráfico."), _
        Array("waterfall", "legend", "none", "Legenda não se aplica a este tipo de gráfico."))
    For i = LBound(cases) To UBound(cases)
        kind = cases(i)(0): key = cases(i)(1): val = cases(i)(2): want = cases(i)(3)
        Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
        If kind = "waterfall" Then      ' uma serie so: a legenda da cascata exige varias
            ws.Range("A1:B1").Value = Array("Etapa", "Valor"): ws.Range("A2:B2").Value = Array("Inicio", 100)
            ws.Range("A3:B3").Value = Array("Fim", "e"): ws.Range("A1:B3").Select
        Else
            FixtureFor(ws, kind).Select
        End If
        Core_Insert kind
        Core_ClearLastNotice
        Core_SetOption key, val
        AssertEqual want, Core_LastNotice(), kind & "/" & key
        AssertEqual 1, Core_NoticeCount(), kind & "/" & key & ": um aviso"
        AssertTrue Core_GetOption(key) <> val, kind & "/" & key & ": nao persiste (" & Core_GetOption(key) & ")"
        Teardown wb
    Next
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Tarefa 15: auditoria do core, divergencia matriz x despacho, anotacao desconhecida, estado real da pizza
' ---------------------------------------------------------------------------------------------------------------------

' Toda funcao publica Core_* de modCoreApi tem "On Error" antes da primeira linha executavel (o loader chama o core por Application.Run: um erro
' nao tratado la viraria o dialogo "Depurar/Finalizar" do VBA na tela do usuario). Le o codigo do proprio projeto (precisa de "Confiar no acesso
' ao modelo de objeto do projeto VBA", como o Test_Ribbon_AllCallbacks_HaveErrorContainment). Excecoes de proposito: getters e setters puros
' (nada que possa falhar) e Core_ApplyPendingLoader (a guarda de reentrada "If mApplyBusy Then Exit Sub" vem antes do On Error; o corpo roda em
' ApplyPendingLoaderImpl, com On Error GoTo Fail).
Public Sub Test_Core_AllPublicApi_HaveErrorContainment()
    Dim cm As Object, lines() As String, i As Long, j As Long, ln As String, nm As String, t As String, n As Long, p As Long, skipList As String
    skipList = ",Core_Ping,Core_Version,Core_LoaderPath,Core_LastNotice,Core_NoticeCount,Core_ClearLastNotice,Core_TestSetApplyBusy,Core_ApplyPendingLoader,"
    Set cm = ThisWorkbook.VBProject.VBComponents("modCoreApi").CodeModule
    lines = Split(Replace(cm.Lines(1, cm.CountOfLines), vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If Left$(ln, 16) = "Public Function " Or Left$(ln, 11) = "Public Sub " Then
            p = InStr(ln, "Core_")
            If p > 0 Then
                nm = Mid$(ln, p, InStr(p, ln & "(", "(") - p)
                If InStr(skipList, "," & nm & ",") = 0 Then
                    For j = i + 1 To UBound(lines)
                        t = Trim$(lines(j))
                        If Len(t) > 0 And Left$(t, 1) <> "'" And Left$(t, 4) <> "Dim " Then
                            AssertTrue Left$(t, 8) = "On Error", "sem On Error antes da 1a linha executavel em " & nm & " (achei: " & t & ")"
                            Exit For
                        End If
                    Next
                    n = n + 1
                End If
            End If
        End If
    Next
    AssertTrue n >= 22, "poucas funcoes Core_ verificadas: " & n
End Sub

' Conteudo do log do build Test (vazio se nao existe).
Private Function TestLogText() As String
    Dim p As String, f As Integer
    p = Errors_LogDir() & "\thinqq.log"
    If Len(Dir(p)) = 0 Then Exit Function
    f = FreeFile
    Open p For Binary Access Read As #f
    TestLogText = Space$(LOF(f))
    Get #f, , TestLogText
    Close #f
End Function

' Matriz e despacho divergem (a matriz aceita o recurso, a familia de despacho nao tem Case): o comando e recusado com o aviso 7002 (e a
' divergencia vai para o log), NADA e gravado nem alterado no grafico; antes caia no padrao da Fase 1 (colunas) e mexia no grafico errado.
' Gancho de teste Core_TestSetFamilyOverride forca uma familia sem despacho. O mesmo comando sem o gancho funciona.
Public Sub Test_Core_DispatchDrift_RefusedWithNotice_NothingChanged()
    Dim wb As Workbook, ws As Worksheet, c As Chart, key As Variant, en As Long, es As String, ed As String, logLen As Long, sTotals As Boolean, sPalette As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    FixtureFor(ws, "stacked").Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    AssertTrue Core_GetToggle("totals") And Core_GetToggle("labels"), "premissa: totais e rotulos ligados"
    sPalette = Core_GetOption("palette")
    logLen = Len(TestLogText())
    On Error GoTo Cleanup
    Core_TestSetFamilyOverride "inexistente"
    For Each key In Array("totals", "labels", "connectors")
        Core_ClearLastNotice
        Core_SetToggle CStr(key), False
        AssertEqual "'" & Kinds_FeatureLabel(CStr(key)) & "' não se aplica a este tipo de gráfico.", Core_LastNotice(), key & ": aviso 7002"
        AssertEqual 1, Core_NoticeCount(), key & ": um aviso"
    Next
    AssertTrue Core_GetToggle("totals") And Core_GetToggle("labels"), "nada foi gravado: totais e rotulos continuam ligados"
    AssertFalse Chart_FindSeries(c, "tqq_total") Is Nothing, "a serie de totais continua no grafico"
    AssertTrue c.SeriesCollection(1).HasDataLabels, "os rotulos continuam ligados"
    Core_ClearLastNotice
    Core_SetOption "palette", "cinza"
    AssertEqual "'Paleta' não se aplica a este tipo de gráfico.", Core_LastNotice(), "palette: aviso 7002"
    AssertEqual sPalette, Core_GetOption("palette"), "palette: recusa antes de gravar"
    AssertTrue InStr(Mid$(TestLogText(), logLen + 1), "despacho sem familia") > 0, "a divergencia vai para o log"
    Core_TestSetFamilyOverride ""
    ' sem o gancho a familia real despacha normalmente
    Core_ClearLastNotice
    Core_SetOption "palette", "cinza"
    AssertEqual "cinza", Core_GetOption("palette"), "sem o gancho a paleta vale"
    Core_SetToggle "totals", False
    AssertFalse Core_GetToggle("totals"), "sem o gancho o toggle vale"
    AssertTrue Chart_FindSeries(c, "tqq_total") Is Nothing, "sem o gancho a serie de totais sai"
    AssertEqual 0, Core_NoticeCount(), "sem avisos"
    Teardown wb
    Exit Sub
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Core_TestSetFamilyOverride ""
    Teardown wb
    Err.Raise en, es, ed
End Sub

' Tipo de anotacao desconhecido (a faixa so envia cagr, diff e bracket): sem aviso ao usuario, sem forma criada, mas a linha vai para o log.
Public Sub Test_Core_AddAnnotation_UnknownKind_LoggedNoNotice()
    Dim wb As Workbook, ws As Worksheet, c As Chart, logLen As Long, shapesBefore As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    FixtureFor(ws, "stacked").Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    shapesBefore = AnnotationShapeCount(c)
    Core_ClearLastNotice
    logLen = Len(TestLogText())
    Core_AddAnnotation "xyz", 1, 2
    AssertEqual 0, Core_NoticeCount(), "sem aviso ao usuario"
    AssertEqual shapesBefore, AnnotationShapeCount(c), "nenhuma forma criada"
    AssertTrue InStr(Mid$(TestLogText(), logLen + 1), "tipo de anotacao desconhecido") > 0, "registrado no log"
    AssertTrue InStr(Mid$(TestLogText(), logLen + 1), "xyz") > 0, "o log traz o nome do tipo"
    Teardown wb
End Sub

' "Alterar dados" na pizza com rotulos fora: o grafico NOVO tem de ter, de verdade, os rotulos fora (posicao, nome da categoria, linhas de
' chamada) e nenhuma legenda; o estado real, nao so a opcao copiada. (No teste de todos os tipos os rotulos ficam desligados e a posicao
' dos rotulos nao e observavel.)
Public Sub Test_Core_RebuildFromRange_Pie_KeepsOutsideLabelsAsRealState()
    Dim wb As Workbook, ws As Worksheet, rng As Range, c As Chart, s As Series
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Set rng = FixtureFor(ws, "pie")
    rng.Select
    Core_Insert "pie"
    Core_SetToggle "pieout", True
    Set c = Core_ActiveChart()
    Set s = c.SeriesCollection(1)
    AssertEqual xlLabelPositionOutsideEnd, s.DataLabels.Position, "premissa: rotulos fora"
    AssertTrue Core_RebuildFromRange(c, rng), "rebuild (" & Core_LastNotice() & ")"
    Set c = Core_ActiveChart()
    Set s = c.SeriesCollection(1)
    AssertTrue s.HasDataLabels, "rotulos ligados"
    AssertEqual xlLabelPositionOutsideEnd, s.DataLabels.Position, "rotulos fora apos Alterar dados"
    AssertTrue s.DataLabels.ShowCategoryName, "nome da categoria nos rotulos fora"
    AssertTrue s.HasLeaderLines, "linhas de chamada"
    AssertFalse c.HasLegend, "sem legenda enquanto os rotulos fora (com os nomes) estao ligados"
    Teardown wb
End Sub

' Tarefa 15b: "Espaco" vale para TODOS os grupos de colunas/barras (uma serie de colunas da combinacao levada ao eixo secundario forma um
' segundo grupo de colunas; antes so ChartGroups(1) recebia), sem tocar nos grupos de linha; e o texto 7002 "nao se aplica" tem uma fonte so
' (Kinds_NotApplicableText), igual ao que o usuario ja via.
Public Sub Test_Core_Gap_EveryColumnGroup_And_NotApplicableTextIsCentral()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, g As ChartGroup, nCol As Long, ct As Long
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("D1:D5").Value = Application.Transpose(Array("Meta", 140, 160, 180, 220))
    ws.Range("A1:D5").Select
    Core_Insert "combo"
    Set cht = Core_ActiveChart()
    AssertEqual "combo", Registry_ChartType(cht)
    cht.SeriesCollection(2).AxisGroup = xlSecondary              ' o usuario leva a 2a serie de colunas ao eixo secundario
    For Each g In cht.ChartGroups
        If g.SeriesCollection(1).ChartType = xlColumnStacked Then nCol = nCol + 1
    Next
    AssertEqual 2, nCol, "pre-condicao: dois grupos de colunas"
    Core_SetOption "gap", "120"
    For Each g In cht.ChartGroups
        ct = g.SeriesCollection(1).ChartType
        If ct = xlColumnStacked Then AssertEqual 120, g.GapWidth, "grupo de colunas do eixo " & g.AxisGroup
    Next
    AssertEqual "120", Registry_GetOption(cht, "gap")
    AssertEqual "Espaço entre colunas não se aplica a este tipo de gráfico.", Kinds_NotApplicableText("gap")
    AssertEqual "Legenda não se aplica a este tipo de gráfico.", Kinds_NotApplicableText("legend")
    AssertEqual "'Rótulos fora da pizza' não se aplica a este tipo de gráfico.", Kinds_NotApplicableText("pieout")
    Core_ClearLastNotice
    Core_SetToggle "ind100", True
    AssertEqual Kinds_NotApplicableText("ind100"), Core_LastNotice(), "toggle recusado pela matriz"
    AssertEqual "'Indicador 100%' não se aplica a este tipo de gráfico.", Core_LastNotice()
    Teardown wb
End Sub
