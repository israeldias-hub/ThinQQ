Attribute VB_Name = "modChartRegistry"
Option Explicit

Public Function Registry_NewId() As String
    Dim chars As String, i As Long, s As String
    chars = "abcdefghijklmnopqrstuvwxyz0123456789"
    Randomize
    For i = 1 To 6
        s = s & Mid$(chars, Int(Rnd * Len(chars)) + 1, 1)
    Next
    Registry_NewId = s
End Function

Public Function Registry_IdFromChart(cht As Chart) As String
    Dim nm As String
    If TypeName(cht.Parent) <> "ChartObject" Then Exit Function
    nm = cht.Parent.Name
    If Left$(nm, Len(ID_PREFIX)) = ID_PREFIX Then Registry_IdFromChart = Mid$(nm, Len(ID_PREFIX) + 1)
End Function

Public Function Registry_IsThinQQChart(cht As Chart) As Boolean
    Registry_IsThinQQChart = Len(Registry_IdFromChart(cht)) > 0
End Function

Public Function Registry_Workbook(cht As Chart) As Workbook
    Set Registry_Workbook = cht.Parent.Parent.Parent   ' ChartObject -> Worksheet -> Workbook
End Function

Public Function Registry_Block(cht As Chart) As Range
    Dim id As String
    id = Registry_IdFromChart(cht)
    If Len(id) > 0 Then Set Registry_Block = CalcSheet_FindBlock(Registry_Workbook(cht), id)
End Function

Public Function Registry_ChartType(cht As Chart) As String
    Dim blk As Range
    Set blk = Registry_Block(cht)
    If Not blk Is Nothing Then Registry_ChartType = CStr(blk.Cells(1, 2).Value)
End Function

Public Function Registry_FindChartObject(wb As Workbook, id As String) As ChartObject
    Dim ws As Worksheet, co As ChartObject
    For Each ws In wb.Worksheets
        Set co = Nothing
        On Error Resume Next
        Set co = ws.ChartObjects(ID_PREFIX & id)
        On Error GoTo 0
        If Not co Is Nothing Then Set Registry_FindChartObject = co: Exit Function
    Next
End Function

Public Function Registry_GetOption(cht As Chart, key As String, Optional defaultValue As String = "") As String
    Dim blk As Range, parts() As String, i As Long, kv() As String
    Registry_GetOption = defaultValue
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    parts = Split(CStr(blk.Cells(1, 4).Value), ";")
    For i = LBound(parts) To UBound(parts)
        kv = Split(parts(i), "=")
        If UBound(kv) = 1 Then
            If kv(0) = key Then Registry_GetOption = kv(1): Exit Function
        End If
    Next
End Function

Public Sub Registry_SetOption(cht As Chart, key As String, value As String)
    Dim blk As Range, parts() As String, i As Long, kv() As String, out As String, found As Boolean
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    parts = Split(CStr(blk.Cells(1, 4).Value), ";")
    For i = LBound(parts) To UBound(parts)
        kv = Split(parts(i), "=")
        If UBound(kv) = 1 Then
            If kv(0) = key Then parts(i) = key & "=" & value: found = True
            out = out & IIf(Len(out) > 0, ";", "") & parts(i)
        End If
    Next
    If Not found Then out = out & IIf(Len(out) > 0, ";", "") & key & "=" & value
    blk.Cells(1, 4).Value = out
End Sub
