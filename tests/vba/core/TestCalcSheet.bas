Attribute VB_Name = "TestCalcSheet"
Option Explicit

Public Sub Test_CalcSheet_Get_CreatesVeryHidden()
    Dim wb As Workbook, ws As Worksheet, active As String
    Set wb = NewTempWorkbook()
    active = wb.ActiveSheet.Name
    Set ws = CalcSheet_Get(wb)
    AssertEqual CALC_SHEET, ws.Name
    AssertEqual xlSheetVeryHidden, ws.Visible
    AssertEqual active, wb.ActiveSheet.Name, "planilha ativa não deve mudar"
    AssertTrue CalcSheet_Get(wb) Is ws, "segunda chamada reutiliza"
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_AllocateAndFind()
    Dim wb As Workbook, blk As Range, found As Range
    Set wb = NewTempWorkbook()
    Set blk = CalcSheet_AllocateBlock(wb, "ab12cd", "stacked", "'Dados'!$A$1:$D$5", 5, 3)
    AssertEqual 6, blk.Rows.Count
    AssertEqual "ab12cd", Block_Meta(blk).Cells(1, 1).Value
    AssertEqual "stacked", Block_Meta(blk).Cells(1, 2).Value
    AssertEqual "'Dados'!$A$1:$D$5", Block_Meta(blk).Cells(1, 3).Value
    AssertEqual 5, Block_Body(blk).Rows.Count
    Set found = CalcSheet_FindBlock(wb, "ab12cd")
    AssertEqual blk.Address, found.Address
    AssertTrue CalcSheet_FindBlock(wb, "zzzzzz") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_TwoBlocks_DoNotOverlap()
    Dim wb As Workbook, b1 As Range, b2 As Range
    Set wb = NewTempWorkbook()
    Set b1 = CalcSheet_AllocateBlock(wb, "aaaaaa", "stacked", "x", 4, 3)
    Set b2 = CalcSheet_AllocateBlock(wb, "bbbbbb", "waterfall", "y", 6, 18)
    AssertTrue b2.Row > b1.Row + b1.Rows.Count, "segundo bloco deve ficar abaixo do primeiro"
    AssertTrue Application.Intersect(b1, b2) Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_DeleteBlock()
    Dim wb As Workbook
    Set wb = NewTempWorkbook()
    CalcSheet_AllocateBlock wb, "aaaaaa", "stacked", "x", 4, 3
    CalcSheet_DeleteBlock wb, "aaaaaa"
    AssertTrue CalcSheet_FindBlock(wb, "aaaaaa") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_Cleanup_RemovesOrphan()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    CalcSheet_AllocateBlock wb, "alive1", "stacked", "x", 4, 3
    Set co = ws.ChartObjects.Add(10, 10, 100, 100)
    co.Name = ID_PREFIX & "alive1"
    AssertEqual 1, CalcSheet_Cleanup(wb)
    AssertTrue CalcSheet_FindBlock(wb, "orphan") Is Nothing
    AssertTrue Not CalcSheet_FindBlock(wb, "alive1") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_Registry_IdAndType()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    CalcSheet_AllocateBlock wb, "k1k1k1", "waterfall", "x", 4, 18
    Set co = ws.ChartObjects.Add(10, 10, 100, 100)
    co.Name = ID_PREFIX & "k1k1k1"
    AssertEqual "k1k1k1", Registry_IdFromChart(co.Chart)
    AssertTrue Registry_IsThinQQChart(co.Chart)
    AssertEqual "waterfall", Registry_ChartType(co.Chart)
    AssertTrue Registry_Workbook(co.Chart) Is wb
    AssertEqual co.Name, Registry_FindChartObject(wb, "k1k1k1").Name
    Set co = ws.ChartObjects.Add(10, 150, 100, 100)
    AssertFalse Registry_IsThinQQChart(co.Chart), "gráfico comum não é ThinQQ"
    CloseTemp wb
End Sub

Public Sub Test_Registry_Options_RoundTrip()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    CalcSheet_AllocateBlock wb, "opt123", "stacked", "x", 4, 3
    Set co = ws.ChartObjects.Add(10, 10, 100, 100)
    co.Name = ID_PREFIX & "opt123"
    AssertEqual "int", Registry_GetOption(co.Chart, "numfmt", "int")
    Registry_SetOption co.Chart, "numfmt", "dec1"
    Registry_SetOption co.Chart, "totals", "1"
    Registry_SetOption co.Chart, "numfmt", "pct0"
    AssertEqual "pct0", Registry_GetOption(co.Chart, "numfmt")
    AssertEqual "1", Registry_GetOption(co.Chart, "totals")
    AssertEqual "numfmt=pct0;totals=1", Registry_Block(co.Chart).Cells(1, 4).Value
    CloseTemp wb
End Sub

Public Sub Test_Registry_NewId_Format()
    Dim id As String, i As Long
    id = Registry_NewId()
    AssertEqual 6, Len(id)
    For i = 1 To 6
        AssertTrue InStr("abcdefghijklmnopqrstuvwxyz0123456789", Mid$(id, i, 1)) > 0
    Next
End Sub

Public Sub Test_Util_ColumnToDoubles()
    Dim wb As Workbook, ws As Worksheet, v() As Double
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:A3").Value = Application.Transpose(Array(1.5, "x", 3))
    v = Util_ColumnToDoubles(ws.Range("A1:A3"))
    AssertEqual 1, LBound(v): AssertEqual 3, UBound(v)
    AssertAlmost 1.5, v(1): AssertAlmost 0, v(2): AssertAlmost 3, v(3)
    v = Util_ColumnToDoubles(ws.Range("A3"))
    AssertEqual 1, UBound(v)
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_SourceAddr_QuoteInSheetName()
    Dim wb As Workbook, blk As Range, src As String
    Set wb = NewTempWorkbook()
    src = "'Plano ''A'' 2026'!$A$1:$D$5"
    Set blk = CalcSheet_AllocateBlock(wb, "q1q1q1", "stacked", src, 4, 3)
    AssertEqual src, CStr(Block_Meta(blk).Cells(1, 3).Value)
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_Cleanup_IgnoresForeignNames()
    Dim wb As Workbook, ws As Worksheet, cs As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    wb.Names.Add Name:="tqq_usuario", RefersTo:="='" & ws.Name & "'!$A$1"
    wb.Names.Add Name:="tqq_const", RefersTo:="=42"
    wb.Names.Add Name:="tqq_bad", RefersTo:="=#REF!"
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    Set cs = CalcSheet_Get(wb)
    cs.Range("Z100").Value = "nao-id"
    wb.Names.Add Name:="tqq_mismatch", RefersTo:="='" & CALC_SHEET & "'!$Z$100"
    ws.Range("A1").Value = "dado"
    AssertEqual 1, CalcSheet_Cleanup(wb)
    AssertTrue NameOk(wb, "tqq_usuario"), "tqq_usuario"
    AssertTrue NameOk(wb, "tqq_const"), "tqq_const"
    AssertTrue NameOk(wb, "tqq_bad"), "tqq_bad"
    AssertTrue NameOk(wb, "tqq_mismatch"), "tqq_mismatch"
    AssertEqual "dado", ws.Range("A1").Value
    AssertEqual "nao-id", cs.Range("Z100").Value
    CloseTemp wb
End Sub

Private Function NameOk(wb As Workbook, nm As String) As Boolean
    Dim n As Name
    For Each n In wb.Names
        If n.Name = nm Then NameOk = True
    Next
End Function

Public Sub Test_Registry_NewId_Unique200()
    Dim wb As Workbook, ids As Object, i As Long, id As String
    Set wb = NewTempWorkbook()
    Set ids = CreateObject("Scripting.Dictionary")
    For i = 1 To 200
        id = Registry_NewId(wb)
        AssertFalse ids.Exists(id), "id repetido: " & id
        ids.Add id, 1
        CalcSheet_AllocateBlock wb, id, "stacked", "x", 1, 1
    Next
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_Allocate_DuplicateId_Raises()
    Dim wb As Workbook, ok As Boolean
    Set wb = NewTempWorkbook()
    CalcSheet_AllocateBlock wb, "dup111", "stacked", "x", 2, 2
    On Error Resume Next
    CalcSheet_AllocateBlock wb, "dup111", "stacked", "y", 2, 2
    ok = (Err.Number <> 0)
    On Error GoTo 0
    AssertTrue ok, "id duplicado deve gerar erro"
    AssertEqual "x", CStr(Block_Meta(CalcSheet_FindBlock(wb, "dup111")).Cells(1, 3).Value)
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_Allocate_BadArgs_Raise()
    Dim wb As Workbook, n1 As Long, n2 As Long
    Set wb = NewTempWorkbook()
    On Error Resume Next
    CalcSheet_AllocateBlock wb, "bad111", "stacked", "x", 0, 2
    n1 = Err.Number: Err.Clear
    CalcSheet_AllocateBlock wb, "bad222", "stacked", "a""b", 2, 2
    n2 = Err.Number: Err.Clear
    On Error GoTo 0
    AssertTrue n1 <> 0, "bodyRows 0"
    AssertTrue n2 <> 0, "aspas"
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_LongSourceAddr_RoundTrips()
    Dim wb As Workbook, src As String, blk As Range
    Set wb = NewTempWorkbook()
    src = "'" & String$(120, "a") & "'!$A$1," & "'" & String$(120, "b") & "'!$A$1," & "'" & String$(120, "c") & "'!$A$1"
    Set blk = CalcSheet_AllocateBlock(wb, "lng111", "stacked", src, 2, 2)
    AssertEqual src, CStr(Block_Meta(blk).Cells(1, 3).Value)
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_Get_ProtectedStructure_RaisesAndRestores()
    Dim wb As Workbook, n As Long, nsheets As Long
    Set wb = NewTempWorkbook()
    nsheets = wb.Worksheets.Count
    wb.Protect Structure:=True
    Application.ScreenUpdating = True
    On Error Resume Next
    CalcSheet_Get wb
    n = Err.Number
    On Error GoTo 0
    wb.Unprotect
    AssertTrue n <> 0, "deve falhar"
    AssertTrue Application.ScreenUpdating, "ScreenUpdating restaurado"
    AssertEqual nsheets, wb.Worksheets.Count
    Application.ScreenUpdating = True
    CloseTemp wb
End Sub

Public Sub Test_Registry_SetOption_Validation_And_Preserve()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, n As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    CalcSheet_AllocateBlock wb, "ops111", "stacked", "x", 2, 2
    Set co = ws.ChartObjects.Add(10, 10, 100, 100)
    co.Name = ID_PREFIX & "ops111"
    Registry_SetOption co.Chart, "a", "1"
    Registry_SetOption co.Chart, "b", "2"
    Registry_SetOption co.Chart, "c", "3"
    On Error Resume Next
    Registry_SetOption co.Chart, "d", "x=y"
    n = Err.Number: Err.Clear
    AssertTrue n <> 0, "valor com ="
    Registry_SetOption co.Chart, "d", "x;y"
    AssertTrue Err.Number <> 0, "valor com ;"
    Err.Clear
    Registry_SetOption co.Chart, "", "x"
    AssertTrue Err.Number <> 0, "chave vazia"
    Err.Clear
    Registry_SetOption co.Chart, "k=", "x"
    AssertTrue Err.Number <> 0, "chave com ="
    On Error GoTo 0
    AssertEqual "a=1;b=2;c=3", Registry_Block(co.Chart).Cells(1, 4).Value
    Registry_SetOption co.Chart, "b", "9"
    AssertEqual "a=1;b=9;c=3", Registry_Block(co.Chart).Cells(1, 4).Value
    Registry_Block(co.Chart).Cells(1, 4).Value = "a=1;lixo;x=p=q;c=3"
    AssertEqual "p=q", Registry_GetOption(co.Chart, "x")
    Registry_SetOption co.Chart, "c", "4"
    AssertEqual "a=1;lixo;x=p=q;c=4", Registry_Block(co.Chart).Cells(1, 4).Value
    CloseTemp wb
End Sub

' ---- Tarefa 0: falha fechada na limpeza ----

Private Sub ResetCleanupHooks()
    Util_FailChartReadForTest = False
    Util_FailChartReadSiteForTest = 0
    Util_LabelBudgetForTest = 0
    Util_LabelReadsForTest = 0
End Sub

' Grafico de colunas embutido com nSeries series sobre Dados!A1:A<nPoints> (preenchida com =ROW()), com ou sem
' rotulos automaticos. A planilha precisa se chamar "Dados".
Private Function PlainChart(ws As Worksheet, nSeries As Long, nPoints As Long, withLabels As Boolean, Optional leftPos As Double = 10) As ChartObject
    Dim co As ChartObject, s As Series, k As Long
    ws.Range("A1").Resize(nPoints, 1).Formula = "=ROW()"
    Set co = ws.ChartObjects.Add(leftPos, 10, 240, 150)
    co.Chart.ChartType = xlColumnClustered
    For k = 1 To nSeries
        Set s = co.Chart.SeriesCollection.NewSeries
        s.Values = "=Dados!$A$1:$A$" & nPoints
        If withLabels Then s.HasDataLabels = True
    Next
    Set PlainChart = co
End Function

' True se ler a propriedade falha (usado para registrar o comportamento real do Excel).
Private Function DataLabelFormulaFails(p As Object) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = p.DataLabel.Formula
    DataLabelFormulaFails = (Err.Number <> 0)
    Err.Clear
End Function

' Falha fechada: se a leitura de qualquer grafico falhar, a limpeza nao remove nada (nem orfaos de verdade)
Public Sub Test_CalcSheet_Cleanup_ReadFailure_DeletesNothing()
    Dim wb As Workbook, ws As Worksheet, n As Long, en As Long, es As String, ed As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    ws.ChartObjects.Add 10, 10, 100, 100                 ' um grafico qualquer: ha leitura a fazer
    On Error GoTo Cleanup
    Util_FailChartReadForTest = True
    n = CalcSheet_Cleanup(wb)
    Util_FailChartReadForTest = False
    AssertEqual 0, n, "com leitura falha nada e removido"
    AssertTrue CalcSheet_CleanupReadFailed(), "a falha fica registrada"
    AssertTrue Not CalcSheet_FindBlock(wb, "orphan") Is Nothing
    AssertEqual 1, CalcSheet_Cleanup(wb), "sem falha o orfao sai"
    AssertFalse CalcSheet_CleanupReadFailed()
    CloseTemp wb
    Exit Sub
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    ResetCleanupHooks
    CloseTemp wb
    Err.Raise en, es, ed
End Sub

' Falha fechada em CADA ponto onde a limpeza le o Excel: forma, item de grupo, planilha de grafico, grafico, serie,
' HasDataLabel e DataLabel.Formula. Em todos: devolve 0, nada e apagado e a falha fica registrada.
Public Sub Test_CalcSheet_Cleanup_ReadFailureAtEverySite_DeletesNothing()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, rect As Shape, site As Long, n As Long
    Dim en As Long, es As String, ed As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    PlainChart ws, 1, 3, True                           ' grafico embutido, com rotulos
    Set co = PlainChart(ws, 1, 3, True, 300)            ' outro, que vai para dentro de um grupo de formas
    Set rect = ws.Shapes.AddShape(msoShapeRectangle, 10, 400, 50, 20)
    ws.Shapes.Range(Array(co.Name, rect.Name)).Group
    wb.Charts.Add                                       ' planilha de grafico
    On Error GoTo Cleanup
    For site = READFAIL_SHAPE To READFAIL_FORMULA
        Util_FailChartReadSiteForTest = site
        n = CalcSheet_Cleanup(wb)
        Util_FailChartReadSiteForTest = 0
        AssertEqual 0, n, "ponto " & site & ": nada removido"
        AssertTrue CalcSheet_CleanupReadFailed(), "ponto " & site & ": falha registrada"
        AssertTrue Not CalcSheet_FindBlock(wb, "orphan") Is Nothing, "ponto " & site & ": orfao intacto"
    Next
    AssertEqual 1, CalcSheet_Cleanup(wb), "sem falha (com forma, grupo, planilha de grafico e rotulos) o orfao sai"
    AssertFalse CalcSheet_CleanupReadFailed()
    CloseTemp wb
    Exit Sub
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    ResetCleanupHooks
    CloseTemp wb
    Err.Raise en, es, ed
End Sub

' Na primeira falha a passada para: os graficos restantes nem sao lidos.
Public Sub Test_CalcSheet_Cleanup_StopsScanningAfterFirstFailure()
    Dim wb As Workbook, ws As Worksheet, n As Long, reads As Long, en As Long, es As String, ed As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    PlainChart ws, 1, 3, True
    PlainChart ws, 1, 3, True, 300
    On Error GoTo Cleanup
    Util_LabelReadsForTest = 0
    Util_FailChartReadSiteForTest = READFAIL_HASLABEL
    n = CalcSheet_Cleanup(wb)
    reads = Util_LabelReadsForTest
    ResetCleanupHooks
    AssertEqual 0, n
    AssertTrue CalcSheet_CleanupReadFailed()
    AssertEqual 1, reads, "so o primeiro ponto do primeiro grafico foi lido"
    CloseTemp wb
    Exit Sub
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    ResetCleanupHooks
    CloseTemp wb
    Err.Raise en, es, ed
End Sub

' Orcamento de pontos por passada: estourou, falha fechada (limite reduzido por gancho de teste).
Public Sub Test_CalcSheet_Cleanup_PointBudgetExceeded_FailsClosed()
    Dim wb As Workbook, ws As Worksheet, n As Long, reads As Long, en As Long, es As String, ed As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3      ' altura 5: ate 5 pontos por serie
    PlainChart ws, 2, 4, True
    On Error GoTo Cleanup
    Util_LabelReadsForTest = 0
    Util_LabelBudgetForTest = 3
    n = CalcSheet_Cleanup(wb)
    reads = Util_LabelReadsForTest
    ResetCleanupHooks
    AssertEqual 0, n, "estourou o orcamento: nada removido"
    AssertTrue CalcSheet_CleanupReadFailed()
    AssertTrue Not CalcSheet_FindBlock(wb, "orphan") Is Nothing
    AssertEqual 3, reads, "para ao estourar o orcamento"
    AssertEqual 1, CalcSheet_Cleanup(wb), "com o orcamento normal o orfao sai"
    CloseTemp wb
    Exit Sub
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    ResetCleanupHooks
    CloseTemp wb
    Err.Raise en, es, ed
End Sub

' Grafico grande (3000 pontos x 10 series, sem rotulos ThinQQ): a varredura de rotulos e limitada pela altura do maior
' bloco (5 linhas aqui), termina rapido e NAO conta como falha (sem o limite, 30000 pontos estourariam o orcamento).
Public Sub Test_CalcSheet_Cleanup_LargeChart_IsBoundedAndNotAFailure()
    Dim wb As Workbook, ws As Worksheet, n As Long, reads As Long, t0 As Double, el As Double
    Dim en As Long, es As String, ed As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    PlainChart ws, 10, 3000, False
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    On Error GoTo Cleanup
    Util_LabelReadsForTest = 0
    t0 = Timer
    n = CalcSheet_Cleanup(wb)
    el = Timer - t0
    reads = Util_LabelReadsForTest
    ResetCleanupHooks
    AssertEqual 1, n, "o orfao sai"
    AssertFalse CalcSheet_CleanupReadFailed(), "grafico grande nao e falha de leitura"
    AssertTrue reads <= 50, "no maximo 5 pontos (altura do bloco) x 10 series; leu " & reads
    AssertTrue el < 10, "a limpeza leva " & Format$(el, "0.00") & " s"
    CloseTemp wb
    Exit Sub
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    ResetCleanupHooks
    CloseTemp wb
    Err.Raise en, es, ed
End Sub

' Serie cuja formula ja cita a planilha interna: nenhum rotulo dela e lido (os rotulos ligam ao mesmo bloco).
Public Sub Test_CalcSheet_Cleanup_SeriesReferencingCalcSheet_ReadsNoLabels()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, s As Series, blk As Range, n As Long, reads As Long
    Dim en As Long, es As String, ed As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    Set blk = CalcSheet_AllocateBlock(wb, "keep", "stacked", "x", 4, 3)
    blk.Cells(2, 2).Resize(4, 1).Formula = "=ROW()"
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    Set co = ws.ChartObjects.Add(10, 10, 240, 150)
    Set s = co.Chart.SeriesCollection.NewSeries
    s.Values = "='" & CALC_SHEET & "'!" & blk.Cells(2, 2).Resize(4, 1).Address(True, True)
    s.HasDataLabels = True
    On Error GoTo Cleanup
    Util_LabelReadsForTest = 0
    n = CalcSheet_Cleanup(wb)
    reads = Util_LabelReadsForTest
    ResetCleanupHooks
    AssertEqual 1, n, "so o orfao sai"
    AssertTrue Not CalcSheet_FindBlock(wb, "keep") Is Nothing, "bloco referenciado pela serie fica"
    AssertEqual 0, reads, "nenhum rotulo lido"
    CloseTemp wb
    Exit Sub
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    ResetCleanupHooks
    CloseTemp wb
    Err.Raise en, es, ed
End Sub

' Bloco referenciado SO por rotulos ligados: fica vivo e a leitura para no primeiro rotulo ligado da serie.
Public Sub Test_CalcSheet_Cleanup_LabelsOnlyReference_KeepsBlock_StopsAtFirstLinkedLabel()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, s As Series, blk As Range, n As Long, reads As Long, i As Long
    Dim en As Long, es As String, ed As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    Set blk = CalcSheet_AllocateBlock(wb, "keep", "stacked", "x", 4, 3)
    blk.Cells(2, 2).Resize(4, 1).Formula = "=ROW()"
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    Set co = PlainChart(ws, 1, 4, True)
    Set s = co.Chart.SeriesCollection(1)
    For i = 1 To 4
        s.Points(i).DataLabel.Formula = "='" & CALC_SHEET & "'!" & blk.Cells(1 + i, 2).Address(True, True)
    Next
    On Error GoTo Cleanup
    Util_LabelReadsForTest = 0
    n = CalcSheet_Cleanup(wb)
    reads = Util_LabelReadsForTest
    ResetCleanupHooks
    AssertEqual 1, n, "so o orfao sai"
    AssertTrue Not CalcSheet_FindBlock(wb, "keep") Is Nothing, "bloco referenciado so por rotulos ligados fica"
    AssertFalse CalcSheet_CleanupReadFailed()
    AssertEqual 1, reads, "para no primeiro rotulo ligado"
    CloseTemp wb
    Exit Sub
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    ResetCleanupHooks
    CloseTemp wb
    Err.Raise en, es, ed
End Sub

' Rotulo de texto comum (nao ligado) nao conta como referencia nem como falha de leitura; nem texto que apenas cita
' "_ThinQQ" sem ser uma referencia (sem "!").
Public Sub Test_CalcSheet_Cleanup_PlainLabels_DoNotBlockCleanup()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, s As Series
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:A3").Value = Application.Transpose(Array("a", "b", "c"))
    ws.Range("B1:B3").Value = Application.Transpose(Array(1, 2, 3))
    Set co = ws.ChartObjects.Add(10, 10, 200, 150)
    Set s = co.Chart.SeriesCollection.NewSeries: s.Values = "=" & ws.Name & "!$B$1:$B$3"
    s.HasDataLabels = True
    s.Points(1).DataLabel.Text = "texto livre"
    s.Points(2).DataLabel.Text = "texto com " & CALC_SHEET & " solto"
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    AssertEqual 1, CalcSheet_Cleanup(wb), "orfao removido apesar de rotulos de texto em outro grafico"
    AssertFalse CalcSheet_CleanupReadFailed()
    CloseTemp wb
End Sub

' Citacao "_ThinQQ!..." que nao pode ser interpretada como endereco: falha fechada, nada e removido.
Public Sub Test_CalcSheet_Cleanup_UnparseableCalcSheetMention_FailsClosed()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, s As Series, n As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    Set co = PlainChart(ws, 1, 3, True)
    Set s = co.Chart.SeriesCollection(1)
    s.Points(1).DataLabel.Text = "ver " & CALC_SHEET & "!zz"          ' texto livre, mas parece uma referencia ilegivel
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    n = CalcSheet_Cleanup(wb)
    AssertEqual 0, n, "mencao ilegivel: nada removido"
    AssertTrue CalcSheet_CleanupReadFailed()
    AssertTrue Not CalcSheet_FindBlock(wb, "orphan") Is Nothing
    CloseTemp wb
End Sub

' Grafico nativo do Excel 2016+ (cascata/chartex): Series.Formula levanta erro 445 e o ThinQQ nunca cria nem copia
' um deles. Nao e falha de leitura: a limpeza continua removendo orfaos na pasta.
Public Sub Test_CalcSheet_Cleanup_NativeChartEx_IsNotAReadFailure()
    Dim wb As Workbook, ws As Worksheet, sh As Shape
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:A4").Value = Application.Transpose(Array("a", "b", "c", "d"))
    ws.Range("B1:B4").Value = Application.Transpose(Array(10, -5, 7, 12))
    ws.Range("A1:B4").Select
    Set sh = ws.Shapes.AddChart2(-1, 119, 10, 10, 300, 200)         ' xlWaterfall (chartex)
    AssertTrue sh.HasChart
    AssertEqual 119, sh.Chart.ChartType
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    AssertEqual 1, CalcSheet_Cleanup(wb), "orfao removido apesar do grafico nativo"
    AssertFalse CalcSheet_CleanupReadFailed()
    CloseTemp wb
End Sub

' Registra o comportamento REAL do Excel (2019, pt-BR) que o codigo de limpeza assume ao ler rotulos.
Public Sub Test_CalcSheet_LabelFormula_ObservedBehavior()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, s As Series, s2 As Series, f As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("B1:B3").Value = Application.Transpose(Array(1, 2, 3))
    Set co = PlainChart(ws, 2, 3, False)
    Set s = co.Chart.SeriesCollection(1): Set s2 = co.Chart.SeriesCollection(2)
    ' (a) sem rotulos: HasDataLabel = False sem erro; ler .DataLabel levanta erro
    AssertFalse s.Points(1).HasDataLabel
    AssertTrue DataLabelFormulaFails(s.Points(1)), "DataLabel de ponto sem rotulo levanta erro (por isso so se le .Formula se HasDataLabel)"
    ' (b) rotulo automatico: .Formula devolve o texto exibido, nao levanta erro e nao comeca com "="
    s.HasDataLabels = True
    AssertTrue s.Points(1).HasDataLabel
    AssertFalse DataLabelFormulaFails(s.Points(1)), "rotulo automatico: .Formula nao levanta erro"
    AssertEqual s.Points(1).DataLabel.Text, s.Points(1).DataLabel.Formula, "rotulo automatico: .Formula = texto exibido"
    ' (c) texto editado a mao: .Formula devolve o texto digitado
    s.Points(1).DataLabel.Text = "texto livre"
    AssertEqual "texto livre", s.Points(1).DataLabel.Formula, "texto editado: .Formula = texto digitado"
    ' (d) rotulo ligado a celula: .Formula devolve a referencia em A1 (mesmo com o estilo L1C1 ligado)
    s.Points(2).DataLabel.Formula = "=Dados!$B$2"
    AssertEqual "=Dados!$B$2", s.Points(2).DataLabel.Formula
    Application.ReferenceStyle = xlR1C1
    On Error Resume Next
    f = s.Points(2).DataLabel.Formula
    Application.ReferenceStyle = xlA1
    On Error GoTo 0
    AssertEqual "=Dados!$B$2", f, "em A1 mesmo com ReferenceStyle = xlR1C1"
    ' (e) rotulo so em alguns pontos / apagado: HasDataLabel False nos demais, sem erro
    s2.Points(2).HasDataLabel = True
    AssertFalse s2.Points(1).HasDataLabel
    AssertTrue s2.Points(2).HasDataLabel
    s2.Points(2).DataLabel.Delete
    AssertFalse s2.Points(2).HasDataLabel
    CloseTemp wb
End Sub

' Tarefa 15b: um rotulo de TEXTO LIVRE que apenas cita "_ThinQQ" (sem ser referencia) antes dos rotulos ligados nao encerra a varredura da serie:
' os rotulos seguintes, ligados ao bloco, ainda contam como referencia e o bloco fica (antes o Exit For no texto livre apagava o bloco).
Public Sub Test_CalcSheet_Cleanup_FreeTextMention_DoesNotStopLabelScan()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, s As Series, blk As Range, n As Long, i As Long
    Dim en As Long, es As String, ed As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    Set blk = CalcSheet_AllocateBlock(wb, "keep", "stacked", "x", 4, 3)
    blk.Cells(2, 2).Resize(4, 1).Formula = "=ROW()"
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    Set co = PlainChart(ws, 1, 4, True)
    Set s = co.Chart.SeriesCollection(1)
    s.Points(1).DataLabel.Text = "nota sobre " & CALC_SHEET & " (texto livre)"
    For i = 2 To 4
        s.Points(i).DataLabel.Formula = "='" & CALC_SHEET & "'!" & blk.Cells(1 + i, 2).Address(True, True)
    Next
    On Error GoTo Cleanup
    n = CalcSheet_Cleanup(wb)
    ResetCleanupHooks
    AssertFalse CalcSheet_CleanupReadFailed()
    AssertTrue Not CalcSheet_FindBlock(wb, "keep") Is Nothing, "bloco ligado pelos rotulos 2..4 fica (o texto livre do rotulo 1 nao encerra a varredura)"
    AssertEqual 1, n, "so o orfao sai"
    CloseTemp wb
    Exit Sub
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    ResetCleanupHooks
    CloseTemp wb
    Err.Raise en, es, ed
End Sub

Public Sub Test_Block_FindCol_Extend_SourceRange()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, blk As Range, blk2 As Range, cols As Long, src As Range, calc As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 10, 5): ws.Range("A3:C3").Value = Array(2023, 20, 6)
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    Set cht = Stacked_Insert(ws, d, "stacked").Chart
    Set blk = Registry_Block(cht)
    cols = blk.Columns.Count
    AssertEqual cols - 1, Block_FindCol(blk, "Pos"): AssertEqual cols, Block_FindCol(blk, "VLine")
    AssertEqual 0, Block_FindCol(blk, "NaoExiste")
    AssertEqual cols - 1, Stacked_PosCol(blk): AssertEqual cols, Stacked_VLineCol(blk)
    Set blk2 = Block_Extend(wb, Registry_IdFromChart(cht), 3)
    AssertEqual cols + 3, blk2.Columns.Count
    AssertEqual cols + 3, Registry_Block(cht).Columns.Count, "o nome tqq_<id> cobre as colunas novas"
    AssertEqual blk.Row, blk2.Row: AssertEqual blk.Rows.Count, blk2.Rows.Count
    AssertEqual cols - 1, Stacked_PosCol(blk2), "Pos continua sendo achada pelo cabecalho depois da extensao"
    AssertEqual cols, Stacked_VLineCol(blk2)
    AssertEqual "", CStr(Block_Body(blk2).Cells(1, cols + 1).Value), "colunas novas vazias"
    Set src = Block_SourceRange(cht)
    AssertEqual "$A$1:$C$3", src.Address(True, True): AssertEqual "Dados", src.Worksheet.Name
    ' extensao recusada quando ha dado a direita (nunca sobrescreve)
    Set calc = wb.Worksheets(CALC_SHEET)
    calc.Cells(blk2.Row + 1, blk2.Column + blk2.Columns.Count).Value = "x"
    On Error Resume Next
    Block_Extend wb, Registry_IdFromChart(cht), 1
    AssertEqual vbObjectError + 6014, Err.Number, "Block_Extend com celula ocupada a direita"
    On Error GoTo 0
    CloseTemp wb
End Sub

Public Sub Test_Area_Cols_ByHeader_AfterExtend()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, blk As Range, p As Long, v As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 10, 5): ws.Range("A3:C3").Value = Array(2023, 20, 6)
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    Set cht = Area_Insert(ws, d, "area").Chart
    Set blk = Registry_Block(cht)
    AssertEqual 2, Area_SeriesCount(blk): p = Area_PosCol(blk): v = Area_VLineCol(blk)
    AssertEqual "Mid1", Block_Body(blk).Cells(1, Area_MidCol(blk, 1)).Value: AssertEqual "Mid2", Block_Body(blk).Cells(1, Area_MidCol(blk, 2)).Value
    AssertEqual "Pos", Block_Body(blk).Cells(1, p).Value: AssertEqual "VLine", Block_Body(blk).Cells(1, v).Value
    Set blk = Block_Extend(wb, Registry_IdFromChart(cht), 5)
    AssertEqual 2, Area_SeriesCount(blk), "numero de series pelo cabecalho Val<k>, nao pela largura"
    AssertEqual p, Area_PosCol(blk): AssertEqual v, Area_VLineCol(blk)
    CloseTemp wb
End Sub

' Onda B (B3.9): cabecalho procurado vazio nao casa a primeira coluna em branco (0); celula de cabecalho com erro (#REF!) nao interrompe a
' busca com erro 13 (a celula e lida de forma defensiva e a busca continua).
Public Sub Test_Block_FindCol_EmptyHeader_And_ErrorCell()
    Dim wb As Workbook, ws As Worksheet, blk As Range
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Set blk = ws.Range("B2:G4")                                  ' linha 1 = meta, linha 2 = cabecalhos
    ws.Range("B3:G3").Value = Array("Categoria", "", "Total", "", "Pos", "VLine")
    ws.Range("E3").Formula = "=NA()"                             ' cabecalho com erro antes de "Pos"
    AssertEqual 0, Block_FindCol(blk, ""), "cabecalho vazio: 0 (nao a primeira coluna em branco)"
    AssertEqual 5, Block_FindCol(blk, "Pos"), "o erro em E3 nao interrompe a busca"
    AssertEqual 6, Block_FindCol(blk, "VLine")
    AssertEqual 3, Block_FindCol(blk, "Total")
    CloseTemp wb
End Sub

' Onda B (B3.9): caminhos de reserva. Bloco legado sem cabecalhos (Fase 1): Pos/VLine pelas duas ultimas colunas e Area_SeriesCount pela
' largura; Block_Extend com id desconhecido levanta 6015; Block_SourceRange devolve Nothing com o endereco da origem invalido.
Public Sub Test_Block_Fallbacks_LegacyBlock_UnknownId_BadSource()
    Dim wb As Workbook, ws As Worksheet, blk As Range, d As clsChartData, cht As Chart, en As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    Set blk = ws.Range("J1:P3")                                  ' 7 colunas sem cabecalhos
    AssertEqual 6, Stacked_PosCol(blk), "legado: Pos = penultima": AssertEqual 7, Stacked_VLineCol(blk), "legado: VLine = ultima"
    AssertEqual 6, Area_PosCol(blk): AssertEqual 7, Area_VLineCol(blk)
    AssertEqual 1, Area_SeriesCount(blk), "legado: (7 - 5) \ 2 series pela largura"
    On Error Resume Next
    Block_Extend wb, "naoexiste", 1
    en = Err.Number
    On Error GoTo 0
    AssertEqual vbObjectError + 6015, en, "Block_Extend com id desconhecido"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 10, 5): ws.Range("A3:C3").Value = Array(2023, 20, 6)
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    Set cht = Stacked_Insert(ws, d, "stacked").Chart
    AssertFalse Block_SourceRange(cht) Is Nothing, "premissa: origem resolvida"
    Registry_Block(cht).Cells(1, 3).Value = "'Planilha que nao existe'!$A$1:$C$3"
    AssertTrue Block_SourceRange(cht) Is Nothing, "planilha inexistente: Nothing"
    Registry_Block(cht).Cells(1, 3).Value = "sem exclamacao"
    AssertTrue Block_SourceRange(cht) Is Nothing, "endereco sem planilha: Nothing"
    CloseTemp wb
End Sub
