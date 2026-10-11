Attribute VB_Name = "TestSamples"
Option Explicit

' Linha do titulo de cada exemplo (1..6) em Exemplos!A.
Private Function TitleRows() As Variant
    TitleRows = Array(1, 22, 43, 64, 85, 106)
End Function

' Linha do titulo de cada secao (7..18) em "Exemplos 2"!A: 1, 22, 43, ..., 232.
Private Function TitleRows2() As Variant
    Dim r(0 To 11) As Long, k As Long
    For k = 0 To 11: r(k) = 1 + 21 * k: Next
    TitleRows2 = r
End Function

' Tamanho do log de erros do build Test (Errors_Log): os exemplos nao devem registrar nada nele.
Private Function ErrLogSize() As Long
    Dim p As String
    p = Errors_LogDir() & "\thinqq.log"
    If Len(Dir(p)) > 0 Then ErrLogSize = FileLen(p)
End Function

' Uma unica pasta de exemplos para todas as verificacoes de conteudo (Samples_Create monta 19 graficos).
Public Sub Test_Samples_BothSheets()
    Dim wb As Workbook, en As Long, es As String, ed As String, logBefore As Long, noticesBefore As Long
    logBefore = ErrLogSize()
    noticesBefore = Core_NoticeCount()
    Set wb = Samples_Create()
    On Error GoTo Cleanup
    AssertEqual "Exemplos", wb.ActiveSheet.Name, "a primeira planilha continua ativa"
    AssertEqual "Exemplos", wb.Worksheets(1).Name
    AssertEqual wb.Worksheets("Exemplos").Index + 1, wb.Worksheets("Exemplos 2").Index, "a segunda planilha vem logo depois da primeira"
    AssertEqual wb.Name, ActiveWorkbook.Name, "a pasta de exemplos fica ativa"
    AssertEqual "Range", TypeName(Selection), "a selecao e uma celula (nao um grafico)"
    AssertEqual "$A$1", Selection.Address, "A1 selecionada na primeira planilha"
    CheckSheet1Create wb
    CheckSheet1Layout wb
    CheckSheet1Content wb
    CheckSheet2KindsAndFeatures wb
    CheckSheet2Layout wb
    CheckSheet2Labels wb
    CheckNoErrorsAnywhere wb
    AssertEqual logBefore, ErrLogSize(), "montar os exemplos nao registra erros no log"
    AssertEqual noticesBefore, Core_NoticeCount(), "montar os exemplos nao mostra avisos ao usuario (em Release seriam MsgBox)"
    CloseTemp wb
    Exit Sub
Cleanup:
    en = Err.Number: es = Err.Source: ed = Err.Description
    CloseTemp wb
    Err.Raise en, es, ed
End Sub

Private Sub CheckSheet1Create(wb As Workbook)
    Dim ws As Worksheet, co As ChartObject, kinds As String
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
End Sub

' Cada grafico fica abaixo do titulo da sua secao e acima do titulo da proxima; graficos nao se sobrepoem.
Private Sub CheckSheet1Layout(wb As Workbook)
    Dim ws As Worksheet, i As Long, j As Long, tr As Variant, a As ChartObject, b As ChartObject
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
End Sub

' O 100% nao tem seta de diferenca (so o indicador 100%, que Ann_ShapeCount nao conta); a cascata que cruza o zero tem totais 'e'.
Private Sub CheckSheet1Content(wb As Workbook)
    Dim ws As Worksheet, co As ChartObject, nWf As Long, found100 As Boolean, crossing As Boolean, body As Range, i As Long
    Dim sr As Series, n As Long, oneSeries As Long, cg As Shape, tl As Series
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
    ' exemplo 1: a linha do CAGR passa acima dos rotulos dos totais da primeira e da ultima coluna
    Set cg = ws.ChartObjects(1).Chart.Shapes("tqq_ann_cagr_1")
    Set tl = Chart_FindSeries(ws.ChartObjects(1).Chart, "tqq_total")
    AssertTrue cg.Top + cg.Height <= tl.Points(1).DataLabel.Top - 8, "CAGR do exemplo 1 sobre o total da 1a coluna"
    AssertTrue cg.Top <= tl.Points(4).DataLabel.Top - 8, "CAGR do exemplo 1 sobre o total da ultima coluna"
    AssertEqual "e", LCase$(ws.Range("B113").Value)
    oneSeries = 0
    For Each co In ws.ChartObjects
        n = 0
        For Each sr In co.Chart.SeriesCollection
            If Not Style_IsHelperSeries(sr) Then n = n + 1
        Next
        If n = 1 Then oneSeries = oneSeries + 1: AssertFalse co.Chart.HasTitle, "grafico de 1 serie sem titulo: " & co.Name
    Next
    AssertTrue oneSeries >= 1, "deve haver ao menos um exemplo de 1 serie"
End Sub

Private Function Ann_ShapeCount(cht As Chart) As Long
    Dim shp As Shape
    For Each shp In cht.Shapes
        ' o indicador 100% (tqq_ann_ind100 e o rotulo) nao e anotacao do usuario: o 100% dos exemplos nasce com ele
        If Left$(shp.Name, Len(ANN_PREFIX)) = ANN_PREFIX And Left$(shp.Name, Len(ANN_PREFIX & "ind100")) <> ANN_PREFIX & "ind100" Then Ann_ShapeCount = Ann_ShapeCount + 1
    Next
End Function

' Existe uma forma de anotacao com este nome no grafico?
Private Function HasShape(cht As Chart, shapeName As String) As Boolean
    Dim shp As Shape
    For Each shp In cht.Shapes
        If shp.Name = shapeName Then HasShape = True: Exit Function
    Next
End Function

' Tipo de cada grafico de "Exemplos 2", na ordem de criacao (a secao 18 tem dois).
Private Function ExpectedKinds2() As Variant
    ExpectedKinds2 = Array("barstacked", "waterfall", "barwaterfall", "mekko", "mekkounits", "gantt", "doughnut", "pie", _
                           "area", "combo", "stacked", "clustered", "clustered")
End Function

Private Sub CheckSheet2KindsAndFeatures(wb As Workbook)
    Dim ws As Worksheet, co As ChartObject, kinds As String, c1 As Chart, c2 As Chart, ek As Variant, i As Long, vl As Series, v As Variant
    Dim wsum As Double, tsum As Double
    Set ws = wb.Worksheets("Exemplos 2")
    AssertEqual 13, ws.ChartObjects.Count, "12 secoes, a ultima com 2 graficos"
    ek = ExpectedKinds2()
    i = 0
    For Each co In ws.ChartObjects
        AssertTrue Registry_IsThinQQChart(co.Chart), co.Name
        kinds = kinds & Registry_ChartType(co.Chart) & ";"
        i = i + 1
        AssertEqual CStr(ek(i - 1)), Registry_ChartType(co.Chart), "tipo do grafico " & i & " de Exemplos 2"
    Next
    AssertTrue InStr(kinds, "barstacked;") > 0 And InStr(kinds, "barwaterfall;") > 0 And InStr(kinds, "mekko;") > 0 And InStr(kinds, "mekkounits;") > 0, kinds
    AssertTrue InStr(kinds, "gantt;") > 0 And InStr(kinds, "pie;") > 0 And InStr(kinds, "doughnut;") > 0 And InStr(kinds, "area;") > 0 And InStr(kinds, "combo;") > 0, kinds
    AssertTrue HasShape(ws.ChartObjects(1).Chart, "tqq_ann_diff_1"), "7: seta de diferenca"
    AssertEqual 2, Waterfall_SeriesCount(ws.ChartObjects(2).Chart), "8: cascata com 2 series"
    AssertTrue HasShape(ws.ChartObjects(4).Chart, "tqq_ann_ind100"), "10: indicador 100% do Mekko (%)"
    AssertTrue Not Chart_FindSeries(ws.ChartObjects(5).Chart, "tqq_vline") Is Nothing, "11: linha de valor"
    ' 11: media PONDERADA pela largura (50*50 + 30*30 + 20*20) / 100 = 38, calculada aqui a partir das celulas
    Set vl = Chart_FindSeries(ws.ChartObjects(5).Chart, "tqq_vline")
    v = vl.Values
    wsum = 0: tsum = 0
    For i = 87 To 89
        wsum = wsum + ws.Cells(i, 2).Value
        tsum = tsum + ws.Cells(i, 2).Value * (ws.Cells(i, 3).Value + ws.Cells(i, 4).Value)
    Next
    AssertAlmost tsum / wsum, CDbl(v(LBound(v))), 0.000001, "11: media ponderada da linha de valor"
    AssertTrue HasShape(ws.ChartObjects(6).Chart, "tqq_ann_bracket_1"), "12: colchete"
    AssertEqual "doughnut", Registry_ChartType(ws.ChartObjects(7).Chart), "13: rosca"
    AssertEqual "pie", Registry_ChartType(ws.ChartObjects(8).Chart), "14: pizza"
    AssertEqual "1", Registry_GetOption(ws.ChartObjects(8).Chart, "pieout"), "14: rotulos fora"
    AssertEqual "1", Registry_GetOption(ws.ChartObjects(10).Chart, "secondary"), "16: linha no eixo secundario"
    AssertTrue Not Chart_FindSeries(ws.ChartObjects(11).Chart, "tqq_vline") Is Nothing, "17: linha de valor"
    ' 17: a linha de valor e a media dos totais das colunas (soma de B213:D216 dividida por 4 anos; calculada aqui, nao pelo bloco)
    Set vl = Chart_FindSeries(ws.ChartObjects(11).Chart, "tqq_vline")
    v = vl.Values
    AssertAlmost Application.WorksheetFunction.Sum(ws.Range("B213:D216")) / 4, CDbl(v(LBound(v))), 0.000001, "17: valor da linha de valor"
    Set c1 = ws.ChartObjects(12).Chart: Set c2 = ws.ChartObjects(13).Chart
    AssertEqual c1.Axes(xlValue).MaximumScale, c2.Axes(xlValue).MaximumScale, "18: mesma escala (maximo)"
    AssertEqual c1.Axes(xlValue).MinimumScale, c2.Axes(xlValue).MinimumScale, "18: mesma escala (minimo)"
    AssertTrue Len(Registry_GetOption(c1, "vmax")) > 0
    AssertEqual ws.ChartObjects(12).Top, ws.ChartObjects(13).Top
End Sub

Private Sub CheckSheet2Layout(wb As Workbook)
    Dim ws As Worksheet, i As Long, j As Long, tr As Variant, a As ChartObject, b As ChartObject, sec As Long
    Set ws = wb.Worksheets("Exemplos 2")
    tr = TitleRows2()
    For i = 1 To ws.ChartObjects.Count
        Set a = ws.ChartObjects(i)
        sec = IIf(i <= 12, i, 12)
        AssertTrue a.Top >= ws.Cells(tr(sec - 1), 1).Top, "grafico " & i & " acima do titulo da secao " & sec
        If sec < 12 Then AssertTrue a.Top + a.Height <= ws.Cells(tr(sec), 1).Top, "grafico " & i & " invade a secao seguinte"
        For j = i + 1 To ws.ChartObjects.Count
            Set b = ws.ChartObjects(j)
            AssertFalse a.Left < b.Left + b.Width And b.Left < a.Left + a.Width And a.Top < b.Top + b.Height And b.Top < a.Top + a.Height, "graficos " & i & " e " & j & " sobrepostos"
        Next
    Next
    For i = 0 To 11
        AssertTrue Len(CStr(ws.Cells(tr(i), 1).Value)) > 0, "titulo " & (i + 7)
    Next
    AssertEqual "7) ", Left$(CStr(ws.Cells(1, 1).Value), 3), "numeracao continua a da primeira planilha"
    AssertEqual "18) ", Left$(CStr(ws.Cells(232, 1).Value), 4)
    ' ColumnWidth depende do DPI e da fonte padrao (o Excel arredonda a largura em pixels): tolerancia de meio caractere
    AssertAlmost 18, ws.Columns("A").ColumnWidth, 0.5, "largura da coluna A"
    AssertAlmost 14, ws.Columns("B").ColumnWidth, 0.5, "largura da coluna B"
End Sub

' Texto realmente exibido em alguns rotulos. As expectativas vem das CELULAS da tabela e da localidade (Format$ usa o separador de milhar, o
' sinal e o percentual do Windows, que o Excel usa por padrao), nunca de literais pt-BR.
Private Sub CheckSheet2Labels(wb As Workbook)
    Dim ws As Worksheet, cht As Chart, txt As String, sub1 As Double, endV As Double, ebitda As Double
    Set ws = wb.Worksheets("Exemplos 2")
    ' 7: totais das barras empilhadas (Loja + Online de Sul e de Norte; tabela A2:C5)
    Set cht = ws.ChartObjects(1).Chart
    AssertEqual Format$(ws.Range("B3").Value + ws.Range("C3").Value, "#,##0"), Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "7: total do Sul"
    AssertEqual Format$(ws.Range("B5").Value + ws.Range("C5").Value, "#,##0"), Chart_FindSeries(cht, "tqq_total").Points(3).DataLabel.Text, "7: total do Norte"
    ' 8: cascata com 2 series (A23:C28): subtotal = Inicio + Ajuste das duas series; fim = subtotal + Queda das duas series
    Set cht = ws.ChartObjects(2).Chart
    sub1 = ws.Range("B24").Value + ws.Range("B25").Value + ws.Range("C24").Value + ws.Range("C25").Value
    endV = sub1 + ws.Range("B27").Value + ws.Range("C27").Value
    AssertEqual Format$(sub1, "#,##0"), Chart_FindSeries(cht, "tqq_lbl").Points(3).DataLabel.Text, "8: subtotal"
    AssertEqual Format$(endV, "#,##0"), Chart_FindSeries(cht, "tqq_lbl").Points(5).DataLabel.Text, "8: fim"
    ' 9: cascata horizontal (A44:B50): receita; EBITDA = receita + custos + despesas; lucro = EBITDA + impostos
    Set cht = ws.ChartObjects(3).Chart
    ebitda = ws.Range("B45").Value + ws.Range("B46").Value + ws.Range("B47").Value
    AssertEqual Format$(ws.Range("B45").Value, "#,##0"), Chart_FindSeries(cht, "tqq_lbl").Points(1).DataLabel.Text, "9: receita"
    AssertEqual Format$(ebitda, "#,##0"), Chart_FindSeries(cht, "tqq_lbl").Points(4).DataLabel.Text, "9: EBITDA"
    AssertEqual Format$(ebitda + ws.Range("B49").Value, "#,##0"), Chart_FindSeries(cht, "tqq_lbl").Points(6).DataLabel.Text, "9: lucro"
    ' 14: pizza com rotulos fora mostra nome e valor da primeira fatia (A149:B152)
    Set cht = ws.ChartObjects(8).Chart
    txt = cht.SeriesCollection(1).Points(1).DataLabel.Text
    AssertTrue InStr(txt, CStr(ws.Range("A150").Value)) > 0 And InStr(txt, Format$(ws.Range("B150").Value, "#,##0")) > 0, "14: rotulo da pizza: " & txt
    ' 16: a margem da combinacao e mostrada como percentual (celulas D192:D194 em 0%)
    Set cht = ws.ChartObjects(10).Chart
    AssertEqual Format$(ws.Range("D192").Value, "0%"), Chart_FindSeries(cht, "Margem").Points(1).DataLabel.Text, "16: margem de 2022"
    AssertEqual Format$(ws.Range("D194").Value, "0%"), Chart_FindSeries(cht, "Margem").Points(3).DataLabel.Text, "16: margem de 2024"
End Sub

' Nenhuma celula das planilhas de exemplos (fora a _ThinQQ, onde NA() e intencional) tem valor de erro (#N/D, #DIV/0!...) e nenhum
' rotulo exibido vira erro.
Private Sub CheckNoErrorsAnywhere(wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name <> CALC_SHEET Then CheckSheetNoErrors ws
    Next
End Sub

Private Sub CheckSheetNoErrors(ws As Worksheet)
    Dim c As Range, co As ChartObject, s As Series, k As Long, t As String
    For Each c In ws.UsedRange.Cells
        AssertFalse IsError(c.Value), ws.Name & "!" & c.Address(False, False) & " tem valor de erro"
    Next
    For Each co In ws.ChartObjects
        For Each s In co.Chart.SeriesCollection
            If s.HasDataLabels Then
                For k = 1 To s.Points.Count
                    If s.Points(k).HasDataLabel Then
                        t = s.Points(k).DataLabel.Text
                        AssertFalse Left$(t, 1) = "#", ws.Name & "/" & co.Name & "/" & s.Name & " ponto " & k & ": rotulo '" & t & "'"
                    End If
                Next
            End If
        Next
    Next
End Sub

' Sem efeito colateral em outra pasta; falha no meio fecha a pasta de exemplos e propaga o erro.
Public Sub Test_Samples_Failure_ClosesWorkbook_NoSideEffects()
    Dim other As Workbook, before As Long, sheetsBefore As Long, errNum As Long, en As Long, es As String, ed As String
    Set other = NewTempWorkbook()
    other.Worksheets(1).Range("A1").Value = "intacto"
    before = Application.Workbooks.Count
    sheetsBefore = other.Sheets.Count
    On Error GoTo Cleanup1                       ' o gancho global nunca pode vazar para os proximos testes
    Util_FailAfterBlockForTest = True
    On Error Resume Next
    Samples_Create
    errNum = Err.Number
    On Error GoTo Cleanup1
    Util_FailAfterBlockForTest = False
    AssertTrue errNum <> 0, "esperava erro"
    AssertEqual before, Application.Workbooks.Count
    AssertEqual sheetsBefore, other.Sheets.Count
    AssertEqual "intacto", other.Worksheets(1).Range("A1").Value
    CloseTemp other
    Exit Sub
Cleanup1:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailAfterBlockForTest = False
    CloseTemp other
    Err.Raise en, es, ed
End Sub

' Core_InsertSamples nunca propaga erro (mesmo com falha) e deixa ScreenUpdating como estava.
Public Sub Test_Core_InsertSamples_NeverRaises()
    Dim before As Long, su As Boolean, en As Long, es As String, ed As String
    su = Application.ScreenUpdating
    before = Application.Workbooks.Count
    On Error GoTo Cleanup2
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
    Exit Sub
Cleanup2:
    en = Err.Number: es = Err.Source: ed = Err.Description
    Util_FailAfterBlockForTest = False
    Watchers_Stop
    Err.Raise en, es, ed
End Sub
