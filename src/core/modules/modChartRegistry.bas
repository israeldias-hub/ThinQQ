Attribute VB_Name = "modChartRegistry"
Option Explicit

Private mSeeded As Boolean

' Id de 6 caracteres [a-z0-9] que não colide com bloco nem gráfico existente em wb (se informado).
Public Function Registry_NewId(Optional wb As Workbook = Nothing) As String
    Dim chars As String, i As Long, s As String, tries As Long
    chars = "abcdefghijklmnopqrstuvwxyz0123456789"
    If Not mSeeded Then Randomize: mSeeded = True
    Do
        s = ""
        For i = 1 To 6
            s = s & Mid$(chars, Int(Rnd * Len(chars)) + 1, 1)
        Next
        tries = tries + 1
        If wb Is Nothing Then Exit Do
        If Not CalcSheet_BlockExists(wb, s) Then
            If Registry_FindChartObject(wb, s) Is Nothing Then Exit Do
        End If
        If tries > 1000 Then Err.Raise vbObjectError + 6005, "Registry_NewId", "Não foi possível gerar um id único."
    Loop
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

' Opções em "chave=valor;chave=valor". Cada par é dividido no PRIMEIRO "=".
Public Function Registry_GetOption(cht As Chart, key As String, Optional defaultValue As String = "") As String
    Dim blk As Range, parts() As String, i As Long, p As Long
    Registry_GetOption = defaultValue
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    parts = Split(CStr(blk.Cells(1, 4).Value), ";")
    For i = LBound(parts) To UBound(parts)
        p = InStr(parts(i), "=")
        If p > 1 Then
            If Left$(parts(i), p - 1) = key Then Registry_GetOption = Mid$(parts(i), p + 1): Exit Function
        End If
    Next
End Function

' Fragmentos desconhecidos/malformados são preservados na reescrita.
Public Sub Registry_SetOption(cht As Chart, key As String, value As String)
    Dim blk As Range, parts() As String, i As Long, p As Long, out As String, found As Boolean
    If Len(key) = 0 Or InStr(key, ";") > 0 Or InStr(key, "=") > 0 Or InStr(value, ";") > 0 Or InStr(value, "=") > 0 Then
        Err.Raise vbObjectError + 6006, "Registry_SetOption", "Chave/valor de opção inválidos (vazio, ';' ou '=' não são permitidos)."
    End If
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    parts = Split(CStr(blk.Cells(1, 4).Value), ";")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            p = InStr(parts(i), "=")
            If p > 1 And Not found Then
                If Left$(parts(i), p - 1) = key Then parts(i) = key & "=" & value: found = True
            End If
            out = out & IIf(Len(out) > 0, ";", "") & parts(i)
        End If
    Next
    If Not found Then out = out & IIf(Len(out) > 0, ";", "") & key & "=" & value
    blk.Cells(1, 4).Value = out
End Sub
