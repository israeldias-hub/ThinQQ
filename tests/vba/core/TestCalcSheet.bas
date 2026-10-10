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
    Dim wb As Workbook
    Set wb = NewTempWorkbook()
    wb.Names.Add Name:="tqq_usuario", RefersTo:="=" & wb.Worksheets(1).Name & "!$A$1"
    wb.Names.Add Name:="tqq_const", RefersTo:="=42"
    wb.Names.Add Name:="tqq_bad", RefersTo:="=#REF!"
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    AssertEqual 1, CalcSheet_Cleanup(wb)
    AssertTrue Not wb.Worksheets(1).Parent.Names("tqq_usuario") Is Nothing
    CloseTemp wb
End Sub
