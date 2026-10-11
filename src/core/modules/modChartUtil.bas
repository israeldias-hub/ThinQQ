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
        If VarType(c.Value) = vbDate Then s = Format$(c.Value, "Short Date") Else s = CStr(v)
    End If
    Util_CellLabel = s
End Function

' SHA-256 (hex minusculo) de um arquivo, em VBA puro (modSha256; sem processo filho). Retorna "" em qualquer falha
' (arquivo ausente/ilegivel): quem chama trata como nao verificado.
Public Function Util_Sha256File(path As String) As String
    On Error GoTo Fail
    Util_Sha256File = Sha256_File(path)
    Exit Function
Fail:
    Util_Sha256File = ""
End Function

' Quantidade de células da referência de categorias (2o argumento de SERIES) da série, lida na fórmula da série:
' reflete linhas inseridas/removidas na origem mesmo quando Series.XValues não mostra a diferença.
' Retorna -1 se o argumento não for uma referência simples a uma planilha da pasta wb.
Public Function Util_SeriesCategoryRefCount(s As Series, wb As Workbook) As Long
    Dim f As String, i As Long, ch As String, inQ As Boolean, inS As Boolean, depth As Long
    Dim argN As Long, startPos As Long, arg As String, bang As Long, sh As String
    On Error GoTo Fail
    Util_SeriesCategoryRefCount = -1
    f = s.Formula
    argN = 1
    startPos = InStr(f, "(") + 1
    If startPos < 2 Then Exit Function
    For i = startPos To Len(f)
        ch = Mid$(f, i, 1)
        If ch = """" And Not inS Then
            inQ = Not inQ
        ElseIf ch = "'" And Not inQ Then
            inS = Not inS
        ElseIf Not inQ And Not inS Then
            If ch = "(" Or ch = "{" Then
                depth = depth + 1
            ElseIf ch = "}" Then
                depth = depth - 1
            ElseIf ch = ")" And depth > 0 Then
                depth = depth - 1
            ElseIf (ch = "," And depth = 0) Or (ch = ")" And depth = 0) Then
                If argN = 2 Then arg = Mid$(f, startPos, i - startPos): Exit For
                argN = argN + 1
                startPos = i + 1
            End If
        End If
    Next
    If Len(arg) = 0 Then Exit Function
    bang = InStrRev(arg, "!")
    If bang < 2 Then Exit Function
    sh = Left$(arg, bang - 1)
    If InStr(sh, "[") > 0 Then Exit Function
    If Left$(sh, 1) = "'" And Right$(sh, 1) = "'" Then sh = Replace(Mid$(sh, 2, Len(sh) - 2), "''", "'")
    Util_SeriesCategoryRefCount = wb.Worksheets(sh).Range(Mid$(arg, bang + 1)).Cells.Count
    Exit Function
Fail:
    Util_SeriesCategoryRefCount = -1
End Function
