Attribute VB_Name = "modChartUtil"
Option Explicit

' Gancho de teste: quando True (e BUILD_CONFIG = "Test"), os Insert falham logo apos alocar o bloco,
' para verificar o rollback. Nunca e usado fora dos testes.
Public Util_FailAfterBlockForTest As Boolean

Public Sub Util_TestHookAfterBlock()
    If BUILD_CONFIG = "Test" Then
        If Util_FailAfterBlockForTest Then Err.Raise vbObjectError + 9999, "ThinQQ", "Falha simulada (teste)"
    End If
End Sub

' Primeira posicao vertical livre (sem sobrepor outros graficos da planilha) para o retangulo proposto.
Public Function Chart_FreeTop(ws As Worksheet, l As Double, t As Double, w As Double, h As Double) As Double
    Dim co As ChartObject, moved As Boolean
    Do
        moved = False
        For Each co In ws.ChartObjects
            If l < co.Left + co.Width And co.Left < l + w And t < co.Top + co.Height And co.Top < t + h Then
                t = co.Top + co.Height + 12
                moved = True
            End If
        Next
    Loop While moved
    Chart_FreeTop = t
End Function

Public Sub Util_EnsureDir(path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub
    If Not fso.FolderExists(fso.GetParentFolderName(path)) Then Util_EnsureDir fso.GetParentFolderName(path)
    fso.CreateFolder path
End Sub

Public Function Util_FileName(path As String) As String
    Util_FileName = Mid$(path, InStrRev(path, "\") + 1)
End Function

Public Function Util_ToDouble(v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then
        If VarType(v) <> vbString Then Util_ToDouble = CDbl(v) Else Util_ToDouble = Val(v)
    End If
End Function

' Coluna de células -> array 1..n de Double (texto/vazio = 0).
Public Function Util_ColumnToDoubles(rng As Range) As Double()
    Dim out() As Double, i As Long, n As Long
    n = rng.Rows.Count
    ReDim out(1 To n)
    For i = 1 To n
        out(i) = Util_ToDouble(rng.Cells(i, 1).Value)
    Next
    Util_ColumnToDoubles = out
End Function

Public Function Chart_FindSeries(cht As Chart, name As String) As Series
    Dim s As Series
    For Each s In cht.SeriesCollection
        If s.Name = name Then Set Chart_FindSeries = s: Exit Function
    Next
End Function

' Texto da célula; evita '#####' de colunas estreitas.
Public Function Util_CellLabel(c As Range) As String
    Dim v As Variant, s As String
    v = c.Value2
    If IsError(v) Then Util_CellLabel = CStr(c.Text): Exit Function
    If VarType(v) = vbString Then Util_CellLabel = v: Exit Function
    s = CStr(c.Text)
    If Left$(s, 1) = "#" Then
        If VarType(c.Value) = vbDate Then s = Format$(c.Value, "dd/mm/yyyy") Else s = CStr(v)
    End If
    Util_CellLabel = s
End Function

' SHA-256 (hex minusculo) de um arquivo via certutil. Retorna "" em qualquer falha (quem chama trata como nao verificado).
' Duplica de proposito o Transport_Sha256File do loader: o core nao depende de codigo do loader ao trocar o loader.
Public Function Util_Sha256File(path As String) As String
    Dim sh As Object, proc As Object, lines() As String, ln As String, i As Long, deadline As Double
    On Error GoTo Fail
    If Len(Dir(path)) = 0 Then Exit Function
    If FileLen(path) = 0 Then Exit Function
    Set sh = CreateObject("WScript.Shell")
    Set proc = sh.Exec("certutil -hashfile """ & path & """ SHA256")
    deadline = CDbl(Now) * 86400 + 60
    Do While proc.Status = 0
        If CDbl(Now) * 86400 > deadline Then proc.Terminate: Exit Function
        DoEvents
    Loop
    lines = Split(Replace(proc.StdOut.ReadAll, vbCr, ""), vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = LCase$(Replace(Trim$(lines(i)), " ", ""))
        If Len(ln) = 64 And Not ln Like "*[!0-9a-f]*" Then Util_Sha256File = ln: Exit Function
    Next
    Exit Function
Fail:
    Util_Sha256File = ""
End Function
