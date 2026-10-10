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
