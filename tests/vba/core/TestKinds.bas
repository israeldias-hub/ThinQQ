Attribute VB_Name = "TestKinds"
Option Explicit

Public Sub Test_Kinds_Valid_And_Family()
    AssertTrue Kinds_IsValid("stacked"): AssertTrue Kinds_IsValid("barwaterfall"): AssertTrue Kinds_IsValid("doughnut")
    AssertFalse Kinds_IsValid("pizza"): AssertFalse Kinds_IsValid("")
    AssertEqual "stacked", Kinds_Family("barclustered")
    AssertEqual "waterfall", Kinds_Family("barwaterfall")
    AssertEqual "area", Kinds_Family("area100")
    AssertEqual "combo", Kinds_Family("combo")
    AssertEqual "mekko", Kinds_Family("mekkounits")
    AssertEqual "gantt", Kinds_Family("gantt")
    AssertEqual "pie", Kinds_Family("doughnut")
    AssertEqual "", Kinds_Family("x")
End Sub

' Uma lista separada por virgula nao e um tipo: a matriz consulta listas com InStr e nao pode aceitar nomes
' vizinhos juntos ("stacked,clustered" casaria com a lista de tipos).
Public Sub Test_Kinds_RejectCommaStrings()
    AssertFalse Kinds_IsValid("stacked,clustered"), "dois tipos validos vizinhos"
    AssertFalse Kinds_IsValid(","): AssertFalse Kinds_IsValid("stacked,"): AssertFalse Kinds_IsValid(",stacked")
    AssertFalse Kinds_IsHorizontal("barstacked,barclustered"): AssertFalse Kinds_IsHorizontal("barstacked,gantt")
    AssertFalse Kinds_Is100("stacked100,barstacked100")
    AssertEqual "", Kinds_Family("stacked,clustered")
    AssertFalse Kinds_Supports("stacked,clustered", "labels")
    AssertFalse Kinds_Supports("stacked", "labels,totals")
    AssertTrue Kinds_Supports("stacked", "labels")
End Sub

Public Sub Test_Kinds_Horizontal_100_And_Refs()
    AssertTrue Kinds_IsHorizontal("barstacked"): AssertTrue Kinds_IsHorizontal("gantt")
    AssertFalse Kinds_IsHorizontal("stacked"): AssertFalse Kinds_IsHorizontal("mekko")
    AssertTrue Kinds_Is100("stacked100"): AssertTrue Kinds_Is100("mekko"): AssertTrue Kinds_Is100("area100")
    AssertFalse Kinds_Is100("mekkounits"): AssertFalse Kinds_Is100("barstacked")
    AssertEqual 20, Kinds_ExpectedCategoryRefs("mekko", 4)
    AssertEqual 4, Kinds_ExpectedCategoryRefs("gantt", 4)
End Sub

Public Sub Test_Kinds_Supports_Matrix()
    AssertTrue Kinds_Supports("stacked", "totals"): AssertFalse Kinds_Supports("clustered", "totals"): AssertTrue Kinds_Supports("mekko", "totals")
    AssertTrue Kinds_Supports("gantt", "labels"): AssertFalse Kinds_Supports("gantt", "numfmt")
    AssertTrue Kinds_Supports("waterfall", "connectors"): AssertFalse Kinds_Supports("line", "connectors")
    AssertTrue Kinds_Supports("gantt", "gap"): AssertFalse Kinds_Supports("pie", "gap"): AssertFalse Kinds_Supports("mekko", "gap")
    AssertTrue Kinds_Supports("stacked100", "ind100"): AssertTrue Kinds_Supports("mekko", "ind100"): AssertFalse Kinds_Supports("mekkounits", "ind100")
    AssertTrue Kinds_Supports("mekkounits", "vline"): AssertFalse Kinds_Supports("stacked100", "vline"): AssertTrue Kinds_Supports("barclustered", "vline")
    AssertTrue Kinds_Supports("combo", "secondary"): AssertFalse Kinds_Supports("stacked", "secondary")
    AssertTrue Kinds_Supports("pie", "pieout"): AssertFalse Kinds_Supports("doughnut", "pieout"): AssertFalse Kinds_Supports("area", "pieout")
    AssertTrue Kinds_Supports("barwaterfall", "cagr"): AssertFalse Kinds_Supports("pie", "cagr"): AssertFalse Kinds_Supports("mekko", "cagr")
    AssertTrue Kinds_Supports("stacked", "diff"): AssertFalse Kinds_Supports("stacked100", "diff"): AssertFalse Kinds_Supports("area100", "diff")
    AssertTrue Kinds_Supports("gantt", "bracket"): AssertFalse Kinds_Supports("stacked", "bracket")
    AssertTrue Kinds_Supports("line", "samescale"): AssertFalse Kinds_Supports("pie", "samescale"): AssertFalse Kinds_Supports("mekko", "samescale")
    AssertTrue Kinds_Supports("stacked", "legend"): AssertFalse Kinds_Supports("gantt", "legend")
    AssertTrue Kinds_Supports("line", "lineend"): AssertTrue Kinds_Supports("combo", "lineend")
    AssertFalse Kinds_Supports("stacked", "lineend"): AssertFalse Kinds_Supports("area", "lineend"): AssertFalse Kinds_Supports("pie", "lineend")
    AssertTrue Kinds_Supports("pie", "palette"): AssertTrue Kinds_Supports("waterfall", "labels")
    AssertFalse Kinds_Supports("stacked", "feature_inexistente")
    AssertFalse Kinds_Supports("tipo_inexistente", "labels")
End Sub

Public Sub Test_Chart_AddHelperSeries_Vertical_IsInvisibleLine()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, body As Range, s As Series
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:B1").Value = Array("Ano", "Valor")
    ws.Range("A2:B2").Value = Array(2022, 10): ws.Range("A3:B3").Value = Array(2023, 20)
    Set d = ChartData_FromRange(ws.Range("A1:B3"))
    Set cht = Stacked_Insert(ws, d, "stacked").Chart
    Set body = Block_Body(Registry_Block(cht))
    Set s = Chart_AddHelperSeries(cht, "tqq_x", body.Cells(2, 2).Resize(2, 1))
    AssertEqual "tqq_x", s.Name
    AssertEqual xlLine, s.ChartType
    AssertEqual msoFalse, s.Format.Line.Visible
    AssertEqual xlMarkerStyleNone, s.MarkerStyle
    AssertEqual xlPrimary, s.AxisGroup
    AssertTrue InStr(s.Formula, CALC_SHEET) > 0
    AssertEqual "='" & CALC_SHEET & "'!" & body.Cells(2, 2).Resize(2, 1).Address(True, True), Block_Ref(body.Cells(2, 2).Resize(2, 1)), "Block_Ref = referencia absoluta na _ThinQQ"
    CloseTemp wb
End Sub

Public Sub Test_Chart_HideHelperLegendEntries_RemovesAllTqq()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, cht As Chart, body As Range, pal As Variant, keys As String, k As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 10, 5): ws.Range("A3:C3").Value = Array(2023, 20, 6)
    Set d = ChartData_FromRange(ws.Range("A1:C3"))
    Set cht = Stacked_Insert(ws, d, "stacked").Chart          ' ja tem tqq_total fora da legenda
    Set body = Block_Body(Registry_Block(cht))
    Chart_AddHelperSeries cht, "tqq_a", body.Cells(2, 2).Resize(2, 1)
    Chart_AddHelperSeries cht, "tqq_b", body.Cells(2, 3).Resize(2, 1)
    AssertTrue cht.Legend.LegendEntries.Count >= 3, "as auxiliares novas aparecem na legenda (estado parcial)"
    Chart_HideHelperLegendEntries cht
    AssertEqual 2, cht.Legend.LegendEntries.Count, "estado parcial resolvido: so Sul e Norte"
    AssertEqual xlLegendPositionRight, cht.Legend.Position, "religar a legenda preserva a posicao"
    AssertEqual 9, cht.Legend.Font.Size, "religar a legenda preserva a fonte"
    cht.HasLegend = False: cht.HasLegend = True               ' religa: todas as entradas voltam
    AssertEqual 5, cht.Legend.LegendEntries.Count
    Chart_HideHelperLegendEntries cht
    AssertEqual 2, cht.Legend.LegendEntries.Count, "so Sul e Norte"
    Chart_HideHelperLegendEntries cht
    AssertEqual 2, cht.Legend.LegendEntries.Count, "idempotente"
    ' mapeamento: as entradas que ficaram sao as das series de dados (cor da chave = cor da serie)
    pal = Style_Palette("padrao")
    For k = 1 To 2: keys = keys & ";" & cht.Legend.LegendEntries(k).LegendKey.Format.Fill.ForeColor.RGB: Next
    keys = keys & ";"
    AssertTrue InStr(keys, ";" & CLng(pal(0)) & ";") > 0 And InStr(keys, ";" & CLng(pal(1)) & ";") > 0, "entradas restantes = Sul e Norte (cores " & keys & ")"
    CloseTemp wb
End Sub

' Colunas empilhadas cru com auxiliares tqq_ NO MESMO grupo das series de dados, intercaladas (S1, tqq_a, S2, tqq_b; caso
' da cascata com varias series): com a legenda a direita o Excel lista a pilha invertida; o mapeamento tem de apagar
' so as entradas tqq_ nos dois casos (legenda a direita e embaixo).
Public Sub Test_Chart_HideHelperLegendEntries_SameGroupStack_ReversedAndNot()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, cht As Chart, s As Series, j As Long, lp As Variant, k As Long, keys As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:E1").Value = Array("Cat", "S1", "tqq_a", "S2", "tqq_b")
    ws.Range("A2:E2").Value = Array("A", 10, 1, 5, 2)
    ws.Range("A3:E3").Value = Array("B", 20, 1, 6, 2)
    For Each lp In Array(xlLegendPositionRight, xlLegendPositionBottom)
        Set co = ws.ChartObjects.Add(10, 10, 460, 280): Set cht = co.Chart
        Do While cht.SeriesCollection.Count > 0: cht.SeriesCollection(1).Delete: Loop
        For j = 1 To 4
            Set s = cht.SeriesCollection.NewSeries
            s.Name = "=Dados!" & ws.Cells(1, 1 + j).Address
            s.Values = "=Dados!" & ws.Range(ws.Cells(2, 1 + j), ws.Cells(3, 1 + j)).Address
            s.XValues = "=Dados!$A$2:$A$3"
        Next
        cht.ChartType = xlColumnStacked
        For j = 1 To 4: cht.SeriesCollection(j).Format.Fill.ForeColor.RGB = RGB(10 * j, 0, 0): Next
        cht.HasLegend = True: cht.Legend.Position = lp
        AssertEqual 4, cht.Legend.LegendEntries.Count
        Chart_HideHelperLegendEntries cht
        AssertEqual 2, cht.Legend.LegendEntries.Count, "so S1 e S2 (legenda " & lp & ")"
        keys = ";"
        For k = 1 To 2: keys = keys & (cht.Legend.LegendEntries(k).LegendKey.Format.Fill.ForeColor.RGB Mod 256) & ";": Next
        AssertTrue InStr(keys, ";10;") > 0 And InStr(keys, ";30;") > 0, "entradas restantes = S1 (10) e S2 (30); chaves " & keys & " (legenda " & lp & ")"
        Chart_HideHelperLegendEntries cht
        AssertEqual 2, cht.Legend.LegendEntries.Count, "idempotente"
        co.Delete
    Next
    CloseTemp wb
End Sub

' Mesmo cenario com a legenda no canto ou arrastada (posicao livre): a Position nao diz a orientacao da lista (medido:
' a arrastada a partir da direita e o canto listam invertido, a arrastada a partir de baixo lista normal); o mapeamento
' forca a legenda a direita so durante o trabalho e a devolve ao lugar. Antes da correcao apagava S1/S2 e deixava as auxiliares.
Public Sub Test_Chart_HideHelperLegendEntries_SameGroupStack_CornerAndCustomLegend()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, cht As Chart, s As Series, j As Long, mode As Variant, k As Long, keys As String
    Dim lft As Double, tp As Double, lp As Long, ct As Variant
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:E1").Value = Array("Cat", "S1", "tqq_a", "S2", "tqq_b")
    ws.Range("A2:E2").Value = Array("A", 10, 1, 5, 2)
    ws.Range("A3:E3").Value = Array("B", 20, 1, 6, 2)
    For Each ct In Array(xlColumnStacked, xlColumnStacked100)
        For Each mode In Array("corner", "custom-from-right", "custom-from-bottom")
            Set co = ws.ChartObjects.Add(10, 10, 460, 280): Set cht = co.Chart
            Do While cht.SeriesCollection.Count > 0: cht.SeriesCollection(1).Delete: Loop
            For j = 1 To 4
                Set s = cht.SeriesCollection.NewSeries
                s.Name = "=Dados!" & ws.Cells(1, 1 + j).Address
                s.Values = "=Dados!" & ws.Range(ws.Cells(2, 1 + j), ws.Cells(3, 1 + j)).Address
                s.XValues = "=Dados!$A$2:$A$3"
            Next
            cht.ChartType = CLng(ct)
            For j = 1 To 4: cht.SeriesCollection(j).Format.Fill.ForeColor.RGB = RGB(10 * j, 0, 0): Next
            cht.HasLegend = True
            Select Case mode
                Case "corner": cht.Legend.Position = xlLegendPositionCorner
                Case "custom-from-right": cht.Legend.Position = xlLegendPositionRight: cht.Legend.Left = cht.Legend.Left - 8: cht.Legend.Top = cht.Legend.Top - 20
                Case Else: cht.Legend.Position = xlLegendPositionBottom: cht.Legend.Left = cht.Legend.Left - 8: cht.Legend.Top = cht.Legend.Top - 20
            End Select
            lp = cht.Legend.Position: lft = cht.Legend.Left: tp = cht.Legend.Top
            Chart_HideHelperLegendEntries cht
            AssertEqual 2, cht.Legend.LegendEntries.Count, mode & ": so S1 e S2"
            keys = ";"
            For k = 1 To 2: keys = keys & (cht.Legend.LegendEntries(k).LegendKey.Format.Fill.ForeColor.RGB Mod 256) & ";": Next
            AssertTrue InStr(keys, ";10;") > 0 And InStr(keys, ";30;") > 0, mode & ": entradas restantes = S1 (10) e S2 (30); chaves " & keys
            AssertEqual lp, cht.Legend.Position, mode & ": posicao da legenda restaurada"
            If mode <> "corner" Then
                AssertAlmost lft, cht.Legend.Left, 1.5, mode & ": Left restaurado"
                AssertAlmost tp, cht.Legend.Top, 1.5, mode & ": Top restaurado"
            End If
            Chart_HideHelperLegendEntries cht
            AssertEqual 2, cht.Legend.LegendEntries.Count, mode & ": idempotente"
            AssertEqual lp, cht.Legend.Position, mode & ": idempotente nao mexe na legenda"
            co.Delete
        Next
    Next
    CloseTemp wb
End Sub

' Barras + auxiliar XY no grupo secundario: o ponto XY cai no fim da barra (posicao medida nos pontos reais,
' sem usar a formula de producao) e depois de mudar os dados continua alinhado; eixo Y secundario 0..n, oculto.
Public Sub Test_Chart_AddHelperSeries_Horizontal_SecondaryXY_AlignedWithBars()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, calc As Worksheet, s As Series, i As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    Set cht = BarsFixture(ws, calc)
    Set s = Chart_AddHelperSeries(cht, "tqq_total", calc.Range("A1:A4"), calc.Range("B1:B4"))
    AssertEqual "tqq_total", s.Name
    AssertEqual xlSecondary, s.AxisGroup
    AssertEqual msoFalse, s.Format.Line.Visible
    AssertEqual xlMarkerStyleNone, s.MarkerStyle
    AssertTrue InStr(s.Formula, CALC_SHEET) > 0
    AssertTrue cht.HasAxis(xlValue, xlSecondary), "eixo Y secundario criado"
    AssertEqual 0, cht.Axes(xlValue, xlSecondary).MinimumScale
    AssertEqual 4, cht.Axes(xlValue, xlSecondary).MaximumScale
    AssertFalse cht.Axes(xlValue, xlSecondary).ReversePlotOrder
    AssertEqual xlTickLabelPositionNone, cht.Axes(xlValue, xlSecondary).TickLabelPosition
    AssertEqual xlTickMarkNone, cht.Axes(xlValue, xlSecondary).MajorTickMark
    AssertEqual msoFalse, cht.Axes(xlValue, xlSecondary).Format.Line.Visible
    AssertFalse cht.Axes(xlValue, xlSecondary).HasMajorGridlines
    AssertTrue cht.Axes(xlCategory, xlPrimary).ReversePlotOrder, "categorias invertidas: primeira no topo"
    For i = 1 To 4: AssertBarEndEqualsXY cht, i, "inicial": Next
    ws.Range("C5").Value = 400: calc.Range("A4").Value = 440      ' Norte da ultima categoria: muda a escala do eixo de valores
    cht.Refresh
    For i = 1 To 4: AssertBarEndEqualsXY cht, i, "apos mudar os dados": Next
    CloseTemp wb
End Sub

Public Sub Test_Chart_HideHelperLegendEntries_BarsWithSecondaryXY_KeepsDataEntries()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, calc As Worksheet, k As Long, keys As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    Set cht = BarsFixture(ws, calc)
    Chart_AddHelperSeries cht, "tqq_total", calc.Range("A1:A4"), calc.Range("B1:B4")
    Chart_AddHelperSeries cht, "tqq_lbl", calc.Range("C1:C4"), calc.Range("B1:B4")
    AssertEqual 4, cht.Legend.LegendEntries.Count, "as auxiliares XY entram na legenda"
    Chart_HideHelperLegendEntries cht
    AssertEqual 2, cht.Legend.LegendEntries.Count
    keys = ";"
    For k = 1 To 2: keys = keys & (cht.Legend.LegendEntries(k).LegendKey.Format.Fill.ForeColor.RGB Mod 256) & ";": Next
    AssertTrue InStr(keys, ";10;") > 0 And InStr(keys, ";20;") > 0, "entradas restantes = as duas series de barras; chaves " & keys
    Chart_HideHelperLegendEntries cht
    AssertEqual 2, cht.Legend.LegendEntries.Count, "idempotente"
    CloseTemp wb
End Sub

' Barras empilhadas cruas (2 series, 4 categorias, categorias invertidas pelo Style_ApplyBase) e, na planilha interna,
' A = total, B = posicao n - i + 0,5, C = valor da segunda serie.
Private Function BarsFixture(ws As Worksheet, ByRef calc As Worksheet) As Chart
    Dim co As ChartObject, cht As Chart, s As Series, j As Long
    ws.Range("A1:C1").Value = Array("Cat", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("A", 10, 5)
    ws.Range("A3:C3").Value = Array("B", 20, 6)
    ws.Range("A4:C4").Value = Array("C", 30, 7)
    ws.Range("A5:C5").Value = Array("D", 40, 8)
    Set co = ws.ChartObjects.Add(10, 10, 460, 280): Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0: cht.SeriesCollection(1).Delete: Loop
    For j = 1 To 2
        Set s = cht.SeriesCollection.NewSeries
        s.Name = "=Dados!" & ws.Cells(1, 1 + j).Address
        s.Values = "=Dados!" & ws.Range(ws.Cells(2, 1 + j), ws.Cells(5, 1 + j)).Address
        s.XValues = "=Dados!$A$2:$A$5"
    Next
    cht.ChartType = xlBarStacked
    Style_ApplyBase cht
    cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(10, 0, 0)
    cht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(20, 0, 0)
    cht.HasLegend = True: cht.Legend.Position = xlLegendPositionRight: cht.Legend.Font.Size = 9
    Set calc = CalcSheet_Get(ws.Parent)
    calc.Range("A1:A4").Value = Application.Transpose(Array(15, 26, 37, 48))
    calc.Range("B1:B4").Value = Application.Transpose(Array(3.5, 2.5, 1.5, 0.5))
    calc.Range("C1:C4").Value = Application.Transpose(Array(5, 6, 7, 8))
    Set BarsFixture = cht
End Function

' O centro do ponto XY da serie auxiliar coincide com o fim da barra da segunda serie (X) e com o centro da faixa (Y).
Private Sub AssertBarEndEqualsXY(cht As Chart, i As Long, ctx As String)
    Dim bar As Point, xy As Point
    Set bar = cht.SeriesCollection(2).Points(i)
    Set xy = Chart_FindSeries(cht, "tqq_total").Points(i)
    AssertAlmost bar.Left + bar.Width, xy.Left, 1.5, ctx & ": X do ponto " & i & " no fim da barra"
    AssertAlmost bar.Top + bar.Height / 2, xy.Top, 1.5, ctx & ": Y do ponto " & i & " no centro da faixa"
End Sub

' Rotulos de auxiliar XY em barras: a direita (>= 0) ou a esquerda (< 0) do ponto; o refresh reaplica pelo sinal da
' coluna TotalPos do bloco e um ponto que nao da acesso ao rotulo (nao plotado, posicao NA(), ou sem rotulo) nao
' interrompe os demais.
Public Sub Test_Labels_Horizontal_RightLeft_And_Refresh_IsolatesUnplottedPoint()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, calc As Worksheet, id As String, blk As Range, body As Range, s As Series
    Dim v As Variant, i As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1): ws.Name = "Dados"
    Set cht = BarsFixture(ws, calc)
    id = Registry_NewId()
    Set blk = CalcSheet_AllocateBlock(wb, id, "barstacked", "Dados!$A$1:$C$5", 5, 4)
    cht.Parent.Name = ID_PREFIX & id
    Set body = Block_Body(blk)                                  ' colunas: 2 Total (X e rotulo), 3 TotalPos (sinal), 4 Pos
    v = Array(15, -26, 37, 48)
    For i = 1 To 4
        body.Cells(i + 1, 1).Value = "C" & i
        body.Cells(i + 1, 2).Value = v(i - 1)
        body.Cells(i + 1, 3).Value = v(i - 1)
        body.Cells(i + 1, 4).Value = 4 - i + 0.5
    Next
    Set s = Chart_AddHelperSeries(cht, "tqq_total", body.Cells(2, 2).Resize(4, 1), body.Cells(2, 4).Resize(4, 1))
    Labels_LinkPoints s, body.Cells(2, 2).Resize(4, 1), "#,##0;-#,##0;", 0, -1, True
    AssertEqual xlLabelPositionRight, s.Points(1).DataLabel.Position
    AssertEqual xlLabelPositionLeft, s.Points(2).DataLabel.Position
    AssertEqual xlLabelPositionRight, s.Points(3).DataLabel.Position
    ' inverte os sinais; o ponto 3 deixa de ser plotado (posicao #N/A) e perde o rotulo (acessar DataLabel levanta erro)
    body.Cells(2, 3).Value = -15: body.Cells(3, 3).Value = 26: body.Cells(5, 3).Value = -48
    body.Cells(4, 4).Formula = "=NA()"
    s.Points(3).HasDataLabel = False
    Labels_RefreshPositions cht
    AssertEqual xlLabelPositionLeft, s.Points(1).DataLabel.Position, "ponto 1 passou a negativo"
    AssertEqual xlLabelPositionRight, s.Points(2).DataLabel.Position, "ponto 2 passou a positivo"
    AssertEqual xlLabelPositionLeft, s.Points(4).DataLabel.Position, "ponto 4 atualizado mesmo com o ponto 3 nao plotado"
    CloseTemp wb
End Sub
