Attribute VB_Name = "TestSamples"
Option Explicit

' Linha do titulo de cada exemplo (1..6) em Exemplos!A.
Private Function TitleRows() As Variant
    TitleRows = Array(1, 22, 43, 64, 85, 106)
End Function

Public Sub Test_Samples_Create()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, kinds As String
    Set wb = Samples_Create()
    Set ws = wb.Worksheets("Exemplos")
    AssertEqual 6, ws.ChartObjects.Count
    For Each co In ws.ChartObjects
        AssertTrue Registry_IsThinQQChart(co.Chart)
        kinds = kinds & Registry_ChartType(co.Chart) & ";"
    Next
    AssertTrue InStr(kinds, "stacked;") > 0 And InStr(kinds, "waterfall;") > 0 And InStr(kinds, "clustered;") > 0 And InStr(kinds, "line;") > 0, kinds
    AssertTrue InStr(kinds, "stacked100;") > 0, kinds
    AssertTrue Not ws.ChartObjects(1).Chart.Shapes("tqq_ann_cagr_1") Is Nothing
    AssertTrue Not ws.ChartObjects(1).Chart.Shapes("tqq_ann_diff_1") Is Nothing
    AssertEqual "e", LCase$(ws.Range("B27").Value)
    CloseTemp wb
End Sub

' Cada grafico fica abaixo do titulo da sua secao e acima do titulo da proxima; graficos nao se sobrepoem.
Public Sub Test_Samples_Layout_NoOverlap()
    Dim wb As Workbook, ws As Worksheet, i As Long, j As Long, tr As Variant, a As ChartObject, b As ChartObject
    Set wb = Samples_Create()
    Set ws = wb.Worksheets("Exemplos")
    tr = TitleRows()
    AssertEqual 6, ws.ChartObjects.Count
    For i = 1 To 6
        Set a = ws.ChartObjects(i)
        AssertTrue a.Top >= ws.Cells(tr(i - 1), 1).Top, "grafico " & i & " acima do titulo"
        If i < 6 Then AssertTrue a.Top + a.Height <= ws.Cells(tr(i), 1).Top, "grafico " & i & " invade a secao seguinte"
        For j = i + 1 To 6
            Set b = ws.ChartObjects(j)
            AssertFalse a.Left < b.Left + b.Width And b.Left < a.Left + a.Width And a.Top < b.Top + b.Height And b.Top < a.Top + a.Height, "graficos " & i & " e " & j & " sobrepostos"
        Next
        AssertTrue Len(CStr(ws.Cells(tr(i - 1), 1).Value)) > 0, "titulo " & i
    Next
    CloseTemp wb
End Sub

' O 100% nao tem seta de diferenca; a cascata que cruza o zero tem totais 'e'.
Public Sub Test_Samples_Content()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, nWf As Long, found100 As Boolean, crossing As Boolean, body As Range, i As Long
    Set wb = Samples_Create()
    Set ws = wb.Worksheets("Exemplos")
    For Each co In ws.ChartObjects
        If Registry_ChartType(co.Chart) = "stacked100" Then
            found100 = True
            AssertEqual 0, Ann_ShapeCount(co.Chart)
        End If
        If Registry_ChartType(co.Chart) = "waterfall" Then
            nWf = nWf + 1
            Set body = Block_Body(Registry_Block(co.Chart))
            For i = 2 To body.Rows.Count
                If body.Cells(i, WF_COL_END).Value < 0 And body.Cells(i, 5).Value > 0 Then crossing = True
            Next
        End If
    Next
    AssertTrue found100, "falta o exemplo 100%"
    AssertEqual 2, nWf
    AssertTrue crossing, "falta a cascata que cruza o zero"
    AssertEqual "e", LCase$(ws.Range("B111").Value)
    AssertEqual "e", LCase$(ws.Range("B113").Value)
    AssertFalse ws.ChartObjects(4).Chart.HasTitle, "grafico de 1 serie sem titulo"
    CloseTemp wb
End Sub

Private Function Ann_ShapeCount(cht As Chart) As Long
    Dim shp As Shape
    For Each shp In cht.Shapes
        If Left$(shp.Name, Len(ANN_PREFIX)) = ANN_PREFIX Then Ann_ShapeCount = Ann_ShapeCount + 1
    Next
End Function

' Sem efeito colateral em outra pasta; falha no meio fecha a pasta de exemplos e propaga o erro.
Public Sub Test_Samples_Failure_ClosesWorkbook_NoSideEffects()
    Dim other As Workbook, before As Long, sheetsBefore As Long, errNum As Long
    Set other = NewTempWorkbook()
    other.Worksheets(1).Range("A1").Value = "intacto"
    before = Application.Workbooks.Count
    sheetsBefore = other.Sheets.Count
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Samples_Create
    errNum = Err.Number
    On Error GoTo 0
    Util_FailAfterBlockForTest = False
    AssertTrue errNum <> 0, "esperava erro"
    AssertEqual before, Application.Workbooks.Count
    AssertEqual sheetsBefore, other.Sheets.Count
    AssertEqual "intacto", other.Worksheets(1).Range("A1").Value
    CloseTemp other
End Sub

' Core_InsertSamples nunca propaga erro (mesmo com falha) e deixa ScreenUpdating como estava.
Public Sub Test_Core_InsertSamples_NeverRaises()
    Dim before As Long, su As Boolean
    su = Application.ScreenUpdating
    before = Application.Workbooks.Count
    Util_FailAfterBlockForTest = True
    Core_InsertSamples
    Util_FailAfterBlockForTest = False
    AssertEqual before, Application.Workbooks.Count
    AssertEqual su, Application.ScreenUpdating
    Core_InsertSamples
    AssertEqual before + 1, Application.Workbooks.Count
    AssertEqual "Exemplos", ActiveSheet.Name
    Watchers_Stop                    ' desfaz os observadores criados por Core_InsertSamples (estado global)
    CloseTemp ActiveWorkbook
End Sub
