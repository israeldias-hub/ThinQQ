Attribute VB_Name = "TestDelta"
Option Explicit
' Parte B (1.5.0): delta entre barras escolhidas com o mouse. Referencias "k:c" (k = ordinal da serie de DADOS; 0 = barra da cascata; na
' comparativa k = lado e c = etapa), extensoes, textos e (B2) geometria independente das formulas de producao.

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

' Sul 10, 20, 30; Norte 5, 6, -4 (categorias T1..T3).
Private Function Fixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Tri", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5)
    ws.Range("A3:C3").Value = Array("T2", 20, 6)
    ws.Range("A4:C4").Value = Array("T3", 30, -4)
    Set Fixture = ws.Range("A1:C4")
End Function

Private Function WaterfallFixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Receita", 100)
    ws.Range("A3:B3").Value = Array("Custo", -30)
    ws.Range("A4:B4").Value = Array("Lucro", "e")
    Set WaterfallFixture = ws.Range("A1:B4")
End Function

Private Function PairFixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Par"
    ws.Range("A1:C1").Value = Array("Etapa", "Seg 1", "Seg 2")
    ws.Range("A2:C2").Value = Array("Bruto", 27, 29)
    ws.Range("A3:C3").Value = Array("Comp 1", -4, -7)
    ws.Range("A4:C4").Value = Array("Líquido", "e", "e")
    Set PairFixture = ws.Range("A1:C4")
End Function

Private Function InsertOn(rng As Range, kind As String) As Chart
    rng.Worksheet.Activate
    rng.Select
    Core_Insert kind
    Set InsertOn = Core_ActiveChart()
End Function

Private Function FirstSeries(c As Chart, wantHelper As Boolean, wantColumn As Boolean) As Long
    Dim i As Long
    For i = 1 To c.SeriesCollection.Count
        If Style_IsHelperSeries(c.SeriesCollection(i)) = wantHelper And Delta_IsColumnSeries(c.SeriesCollection(i)) = wantColumn Then FirstSeries = i: Exit Function
    Next
End Function

Public Sub Test_Delta_Text()
    Dim f As String
    f = "0"
    AssertEqual "+5 (+50%)", Delta_Text(10, 15, "both", f)
    AssertEqual "+5", Delta_Text(10, 15, "abs", f)
    AssertEqual "+50%", Delta_Text(10, 15, "pct", f)
    AssertEqual "-14 (-140%)", Delta_Text(10, -4, "both", f)
    AssertEqual "+5", Delta_Text(0, 5, "both", f), "origem zero: so o valor (como a Seta diferenca)"
    AssertEqual "n/d", Delta_Text(0, 5, "pct", f)
    AssertEqual "-75%", Delta_Text(-4, -7, "pct", f), "(B - A) / |A|, como a Seta diferenca (a seta de variacao da comparativa daria +75%)"
    AssertEqual "0 (0%)", Delta_Text(5, 5, "both", f), "sem variacao: sem sinal"
    AssertEqual "+3 (+75%)", Delta_Text(-4, -1, "", f), "modo desconhecido = ambos; origem negativa: (B - A) / |A|"
    f = Style_NumberFormat("int")
    AssertEqual Format$(1500, "+" & Split(f & ";", ";")(0) & ";-" & Split(f & ";", ";")(0) & ";0") & " (+150%)", Delta_Text(1000, 2500, "both", f), "formato do grafico com a localidade"
    AssertEqual "abs", Delta_NextFormat("both"): AssertEqual "pct", Delta_NextFormat("abs"): AssertEqual "both", Delta_NextFormat("pct"): AssertEqual "abs", Delta_NextFormat("")
End Sub

Public Sub Test_Delta_ParseRef()
    Dim k As Long, c As Long
    AssertTrue Delta_ParseRef("2:3", k, c): AssertEqual 2, k: AssertEqual 3, c
    AssertTrue Delta_ParseRef("0:1", k, c), "barra da cascata"
    AssertFalse Delta_ParseRef("2", k, c): AssertFalse Delta_ParseRef("a:1", k, c): AssertFalse Delta_ParseRef("1:0", k, c): AssertFalse Delta_ParseRef("-1:2", k, c)
    AssertFalse Delta_ParseRef("1e2:1", k, c), "notacao cientifica": AssertFalse Delta_ParseRef("&H1:1", k, c), "hexadecimal"
    AssertFalse Delta_ParseRef("99999999999:1", k, c), "grande demais para Long (sem erro)": AssertFalse Delta_ParseRef("1:2:3", k, c)
    AssertFalse Delta_ParseRef("", k, c): AssertFalse Delta_ParseRef(" 1:2", k, c): AssertFalse Delta_ParseRef("1,5:2", k, c)
End Sub

Public Sub Test_Delta_ResolvePick_PerFamily()
    Dim wb As Workbook, c As Chart, ref As String, s As Long
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    AssertTrue Delta_ResolvePick(c, Delta_SeriesIndex(c, 2), 3, ref): AssertEqual "2:3", ref
    AssertFalse Delta_ResolvePick(c, Delta_SeriesIndex(c, 1), 4, ref), "categoria inexistente"
    s = FirstSeries(c, True, False)
    If s > 0 Then AssertFalse Delta_ResolvePick(c, s, 1, ref), "auxiliar tqq_* nao e barra"
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "stacked")
    Core_SetToggle "totals", True
    s = Delta_FindSeriesByName(c, "tqq_total")
    AssertTrue s > 0, "a empilhada com Totais tem a auxiliar tqq_total"
    AssertFalse Delta_ResolvePick(c, s, 1, ref), "auxiliar tqq_total nao e barra"
    AssertEqual 2, Delta_DataSeriesCount(c), "auxiliares nao contam"
    AssertTrue Delta_ResolvePick(c, Delta_SeriesIndex(c, 2), 1, ref): AssertEqual "2:1", ref
    AssertEqual 2, Delta_DataOrdinal(c, Delta_SeriesIndex(c, 2))
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "stacked100")
    AssertFalse Delta_ResolvePick(c, Delta_SeriesIndex(c, 1), 1, ref), "100% nao tem o recurso (matriz)"
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(WaterfallFixture(wb), "waterfall")
    AssertTrue Delta_ResolvePick(c, FirstSeries(c, True, True), 2, ref): AssertEqual "0:2", ref, "cascata: a barra da categoria"
    AssertFalse Delta_ResolvePick(c, FirstSeries(c, True, True), 4, ref)
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(PairFixture(wb), "waterfallpair")
    s = FirstSeries(c, True, True)
    If s = 0 Then s = FirstSeries(c, False, True)
    AssertTrue Delta_ResolvePick(c, s, 1, ref): AssertEqual "1:1", ref, "fatia A da etapa 1"
    AssertTrue Delta_ResolvePick(c, s, 2, ref): AssertEqual "2:1", ref, "fatia B da etapa 1"
    AssertFalse Delta_ResolvePick(c, s, 3, ref), "vao entre etapas"
    AssertTrue Delta_ResolvePick(c, s, 4, ref): AssertEqual "1:2", ref
    If Delta_FindSeriesByName(c, "tqq_pd") = 0 Then Wfp_SetArrows c, True
    AssertTrue Delta_FindSeriesByName(c, "tqq_pd") > 0, "setas de variacao ligadas"
    AssertFalse Delta_ResolvePick(c, Delta_FindSeriesByName(c, "tqq_pd"), 1, ref), "seta de variacao (XY) nao e barra"
    AssertFalse Delta_ResolvePick(c, s, 6, ref), "vao entre etapas 2 e 3"
    AssertTrue Delta_ResolvePick(c, s, 8, ref): AssertEqual "2:3", ref, "fatia B da ultima etapa (3n - 1 = 8)"
    AssertFalse Delta_ResolvePick(c, s, 9, ref), "alem da ultima fatia"
    Teardown wb
End Sub

Public Sub Test_Delta_Extent_PerFamily()
    Dim wb As Workbook, c As Chart, v As Double, lo As Double, hi As Double, ref As String
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "stacked")
    AssertTrue Delta_Extent(c, "2:1", v, lo, hi): AssertEqual 5, v: AssertEqual 10, lo: AssertEqual 15, hi
    AssertTrue Delta_Extent(c, "2:3", v, lo, hi): AssertEqual -4, v: AssertEqual -4, lo: AssertEqual 0, hi, "segmento negativo desce do zero"
    AssertTrue Delta_Extent(c, "1:3", v, lo, hi): AssertEqual 30, v: AssertEqual 0, lo: AssertEqual 30, hi
    AssertFalse Delta_Extent(c, "3:1", v, lo, hi), "serie inexistente"
    AssertEqual 26, Delta_SpanEdge(c, 1, 2, True), "maior pilha positiva de T1..T2 (20 + 6)"
    AssertEqual -4, Delta_SpanEdge(c, 3, 3, False)
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    AssertTrue Delta_Extent(c, "2:3", v, lo, hi): AssertEqual -4, lo: AssertEqual 0, hi
    AssertEqual 30, Delta_SpanEdge(c, 1, 3, True)
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "line")
    AssertTrue Delta_Extent(c, "2:3", v, lo, hi): AssertEqual -4, lo: AssertEqual -4, hi, "linha: o proprio ponto"
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "combo")
    AssertTrue Delta_Extent(c, "1:3", v, lo, hi): AssertEqual 30, v: AssertEqual 0, lo: AssertEqual 30, hi, "combinacao: coluna empilhada"
    AssertTrue Delta_Extent(c, "2:3", v, lo, hi): AssertEqual -4, v: AssertEqual -4, lo: AssertEqual -4, hi, "combinacao: a linha e o proprio ponto"
    AssertTrue Delta_ResolvePick(c, Delta_SeriesIndex(c, 2), 2, ref): AssertEqual "2:2", ref, "combinacao: a linha tambem e escolhida"
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(WaterfallFixture(wb), "barwaterfall")
    AssertTrue Delta_Extent(c, "0:1", v, lo, hi): AssertEqual 100, v: AssertEqual 0, lo: AssertEqual 100, hi, "cascata horizontal: o mesmo bloco"
    AssertFalse Delta_Extent(c, "1:1", v, lo, hi), "na cascata k e sempre 0"
    AssertFalse Delta_Extent(c, "0:4", v, lo, hi), "categoria inexistente"
    AssertEqual 100, Delta_SpanEdge(c, 1, 3, True)
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(WaterfallFixture(wb), "waterfall")
    AssertTrue Delta_Extent(c, "0:2", v, lo, hi): AssertEqual 70, v, "valor acumulado (End)": AssertEqual 70, lo: AssertEqual 100, hi
    AssertTrue Delta_Extent(c, "0:3", v, lo, hi): AssertEqual 70, v: AssertEqual 0, lo, "total vai do zero"
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = InsertOn(PairFixture(wb), "waterfallpair")
    AssertTrue Delta_Extent(c, "2:2", v, lo, hi): AssertEqual -7, v, "Val do lado B na etapa 2": AssertEqual 22, lo: AssertEqual 29, hi
    AssertEqual 5, Delta_AxisIndex(c, "2:2"), "fatia B da etapa 2 = 3*(2-1)+2"
    AssertFalse Delta_Extent(c, "3:1", v, lo, hi)
    Teardown wb
End Sub

' ---------- B2: a seta (geometria independente: regra de tres propria com PlotArea.Inside* e a escala reais) ----------

' Regra de tres propria no grafico (independente de Geom_*): y de um valor nas colunas, x nas barras.
Private Function YOf(c As Chart, v As Double) As Double
    YOf = c.PlotArea.InsideTop + c.PlotArea.InsideHeight * (c.Axes(xlValue).MaximumScale - v) / (c.Axes(xlValue).MaximumScale - c.Axes(xlValue).MinimumScale)
End Function

' A mesma regra no eixo de valores SECUNDARIO (a linha da combinacao).
Private Function YSecOf(c As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = c.Axes(xlValue, xlSecondary)
    YSecOf = c.PlotArea.InsideTop + c.PlotArea.InsideHeight * (ax.MaximumScale - v) / (ax.MaximumScale - ax.MinimumScale)
End Function

Private Function XOf(c As Chart, v As Double) As Double
    XOf = c.PlotArea.InsideLeft + c.PlotArea.InsideWidth * (v - c.Axes(xlValue).MinimumScale) / (c.Axes(xlValue).MaximumScale - c.Axes(xlValue).MinimumScale)
End Function

Private Function HasShape(c As Chart, nm As String) As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = c.Shapes(nm)
    HasShape = Not s Is Nothing
End Function

Private Function Relayout(c As Chart) As Chart
    ForceLayout c
    Ann_Refresh c
    ForceLayout c
    Set Relayout = c
End Function

Private Function Overlaps(a As Shape, b As Shape) As Boolean
    Overlaps = Not (a.Left + a.Width <= b.Left Or b.Left + b.Width <= a.Left Or a.Top + a.Height <= b.Top Or b.Top + b.Height <= a.Top)
End Function

' Review Focus 4: 1a serie T1 (10) -> 2a serie T3 (-4) em colunas agrupadas: perna da origem 3 pt acima da barra, ponta no destino 3 pt
' acima do zero (barra negativa), linha principal acima de TODAS as barras do trecho (a maior e 30) e do rotulo, texto "-14 (-140%)".
Public Sub Test_Delta_Columns_ElbowGeometry()
    Dim wb As Workbook, c As Chart, t1 As Shape, t2 As Shape, ln As Shape, lbl As Shape, pa As Point, pb As Point, xA As Double, xB As Double
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    Core_ClearLastNotice
    Core_AddDelta "1:1", "2:3"
    AssertEqual "", Core_LastNotice()
    AssertEqual "kind=delta;v=1;from=1:1;to=2:3;fmt=both", c.Shapes("tqq_ann_delta_1").AlternativeText
    Relayout c
    Set ln = c.Shapes("tqq_ann_delta_1"): Set t1 = c.Shapes("tqq_ann_delta_1_t1"): Set t2 = c.Shapes("tqq_ann_delta_1_t2"): Set lbl = c.Shapes("tqq_ann_delta_1_lbl")
    Set pa = c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(1): Set pb = c.SeriesCollection(Delta_SeriesIndex(c, 2)).Points(3)
    xA = pa.Left + pa.Width / 2: xB = pb.Left + pb.Width / 2
    AssertAlmost xA, t1.Left, 1, "perna da origem no centro da barra Sul T1"
    AssertAlmost YOf(c, 10) - 3, t1.Top + t1.Height, 1, "perna da origem comeca 3 pt acima da barra"
    AssertAlmost xB, t2.Left, 1, "perna do destino no centro da barra Norte T3"
    AssertAlmost YOf(c, 0) - 3, t2.Top + t2.Height, 1, "ponta no destino: 3 pt acima do zero (barra negativa)"
    AssertEqual CLng(msoArrowheadTriangle), CLng(t2.Line.EndArrowheadStyle)
    AssertTrue (t2.VerticalFlip = msoFalse), "a ponta (fim da linha) fica embaixo, no destino"
    AssertAlmost ln.Top, t1.Top, 0.5, "linha principal na altura do topo das pernas"
    AssertTrue lbl.Top + lbl.Height <= YOf(c, 30) - 8 + 0.5, "rotulo acima da maior barra do trecho (30) com folga"
    AssertAlmost (xA + xB) / 2, lbl.Left + lbl.Width / 2, 1, "rotulo centrado"
    AssertEqual Delta_Text(10, -4, "both", Style_NumberFormat(Registry_GetOption(c, "numfmt", "int"))), CStr(lbl.TextFrame2.TextRange.Text)
    Teardown wb
End Sub

' Duas setas com trechos sobrepostos: a mais estreita fica mais perto das barras e a larga sobe ate os rotulos nao se tocarem.
Public Sub Test_Delta_Lanes_NestedArrows()
    Dim wb As Workbook, c As Chart, l1 As Shape, l2 As Shape
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    Core_AddDelta "1:1", "1:3"          ' larga (T1 -> T3)
    Core_AddDelta "1:2", "2:2"          ' estreita (dentro de T2)
    Relayout c
    Set l1 = c.Shapes("tqq_ann_delta_1_lbl"): Set l2 = c.Shapes("tqq_ann_delta_2_lbl")
    AssertTrue c.Shapes("tqq_ann_delta_2").Top > c.Shapes("tqq_ann_delta_1").Top, "a estreita fica mais baixa (perto das barras)"
    AssertTrue c.Shapes("tqq_ann_delta_2").Top - c.Shapes("tqq_ann_delta_1").Top >= l1.Height / 2 + l2.Height / 2 + 6 - 0.5, "faixas separadas"
    AssertFalse Overlaps(l1, l2), "rotulos nao se tocam"
    Teardown wb
End Sub

' Mesma borda (dados iguais nas categorias): sem as faixas as duas setas ficariam na mesma altura; com elas a mais larga sobe o passo exato
' (meia altura de cada rotulo + 6 pt) e os rotulos nao se tocam.
Public Sub Test_Delta_Lanes_SameEdge_StepApart()
    Dim wb As Workbook, c As Chart, rng As Range, l1 As Shape, l2 As Shape
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Worksheet.Range("B2:B4").Value = Application.Transpose(Array(30, 30, 30))
    Set c = InsertOn(rng, "clustered")
    Core_AddDelta "1:1", "1:3"
    Core_AddDelta "1:2", "2:3"
    Relayout c
    Set l1 = c.Shapes("tqq_ann_delta_1_lbl"): Set l2 = c.Shapes("tqq_ann_delta_2_lbl")
    AssertAlmost l1.Height / 2 + l2.Height / 2 + 6, c.Shapes("tqq_ann_delta_2").Top - c.Shapes("tqq_ann_delta_1").Top, 0.6, "passo exato entre as faixas"
    AssertFalse Overlaps(l1, l2), "rotulos nao se tocam"
    Teardown wb
End Sub

' Barras horizontais: tudo girado (pernas horizontais, linha principal vertical a direita das barras).
Public Sub Test_Delta_Bars_Horizontal()
    Dim wb As Workbook, c As Chart, t1 As Shape, ln As Shape
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "barclustered")
    Core_AddDelta "1:1", "1:3"
    Relayout c
    Set t1 = c.Shapes("tqq_ann_delta_1_t1"): Set ln = c.Shapes("tqq_ann_delta_1")
    AssertAlmost XOf(c, 10) + 3, t1.Left, 1, "perna da origem comeca 3 pt depois do fim da barra"
    AssertTrue ln.Width < 0.5, "linha principal vertical"
    AssertTrue ln.Left >= XOf(c, 30) + 8 - 0.5, "a direita da maior barra do trecho"
    AssertTrue c.Shapes("tqq_ann_delta_1_lbl").Left + c.Shapes("tqq_ann_delta_1_lbl").Width <= Geom_ChartRightEdge(c) + 0.5, "rotulo dentro da borda livre"
    Teardown wb
End Sub

' Duas barras inteiramente negativas: cotovelo ABAIXO.
Public Sub Test_Delta_NegativeBelow()
    Dim wb As Workbook, c As Chart, rng As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Worksheet.Range("B2").Value = -5: rng.Worksheet.Range("B3").Value = -8
    Set c = InsertOn(rng, "clustered")
    Core_AddDelta "1:1", "1:2"
    Relayout c
    AssertTrue c.Shapes("tqq_ann_delta_1").Top >= YOf(c, -8) + 8 - 0.5, "linha principal abaixo da barra mais funda"
    AssertAlmost YOf(c, -5) + 3, c.Shapes("tqq_ann_delta_1_t1").Top, 1, "perna da origem 3 pt abaixo da barra"
    Teardown wb
End Sub

' Escala fixa no maior valor: sem sala, a area de plotagem desce e o rotulo fica dentro do grafico.
Public Sub Test_Delta_RoomTop_FixedMax()
    Dim wb As Workbook, c As Chart, top0 As Double
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    c.Axes(xlValue).MaximumScale = 30
    ForceLayout c
    top0 = c.PlotArea.InsideTop
    Core_AddDelta "1:1", "1:3"
    Relayout c
    AssertTrue c.Shapes("tqq_ann_delta_1_lbl").Top >= 1.5, "rotulo inteiro dentro do grafico (" & c.Shapes("tqq_ann_delta_1_lbl").Top & ")"
    AssertTrue c.PlotArea.InsideTop > top0, "a area de plotagem desceu"
    AssertTrue c.Shapes("tqq_ann_delta_1_lbl").Top + c.Shapes("tqq_ann_delta_1_lbl").Height <= YOf(c, 30) - 8 + 0.5, "e o rotulo continua acima da barra de 30"
    Teardown wb
End Sub

' Dado novo: texto e geometria acompanham (sem refazer a seta).
Public Sub Test_Delta_DataChange_TextAndGeometry()
    Dim wb As Workbook, c As Chart, rng As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    Set c = InsertOn(rng, "clustered")
    Core_AddDelta "1:1", "1:3"
    rng.Worksheet.Range("B4").Value = 40
    Application.Calculate
    Relayout c
    AssertEqual Delta_Text(10, 40, "both", Style_NumberFormat(Registry_GetOption(c, "numfmt", "int"))), CStr(c.Shapes("tqq_ann_delta_1_lbl").TextFrame2.TextRange.Text)
    AssertTrue c.Shapes("tqq_ann_delta_1_lbl").Top + c.Shapes("tqq_ann_delta_1_lbl").Height <= YOf(c, 40) - 8 + 0.5
    Teardown wb
End Sub

' Referencia que nao existe mais: seta oculta (nao apagada).
Public Sub Test_Delta_InvalidRefHidden()
    Dim wb As Workbook, c As Chart
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    Core_AddDelta "1:1", "1:3"
    c.Shapes("tqq_ann_delta_1").AlternativeText = "kind=delta;v=1;from=1:1;to=1:9;fmt=both"
    Ann_Refresh c
    AssertEqual CLng(msoFalse), CLng(c.Shapes("tqq_ann_delta_1").Visible)
    AssertEqual CLng(msoFalse), CLng(c.Shapes("tqq_ann_delta_1_lbl").Visible)
    AssertEqual CLng(msoFalse), CLng(c.Shapes("tqq_ann_delta_1_t2").Visible)
    Teardown wb
End Sub

' Recusas: tipo sem o recurso (7001) e mesma barra / referencia invalida (7003) viram aviso; nada criado.
Public Sub Test_Delta_Refusals()
    Dim wb As Workbook, c As Chart, ws As Worksheet
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    AssertEqual "", Ann_RefusalText(c, "delta"), "tipo com o recurso: sem texto de recusa"
    Core_ClearLastNotice
    Core_AddDelta "1:2", "1:2"
    AssertEqual DELTA_BAD_PICK, Core_LastNotice()
    Core_AddDelta "1:2", "5:1"
    AssertEqual DELTA_BAD_PICK, Core_LastNotice()
    AssertFalse HasShape(c, "tqq_ann_delta_1")
    Teardown wb
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = "Cat": ws.Range("B1").Value = "V": ws.Range("A2").Value = "a": ws.Range("B2").Value = 1: ws.Range("A3").Value = "b": ws.Range("B3").Value = 2
    Set c = InsertOn(ws.Range("A1:B3"), "pie")
    Core_ClearLastNotice
    Core_AddDelta "1:1", "1:2"
    AssertEqual ANN_DELTA_REFUSAL, Core_LastNotice()
    Teardown wb
End Sub

' Falha no meio da criacao (gancho): nenhuma forma da seta sobra.
Public Sub Test_Delta_AddFailure_RollsBack()
    Dim wb As Workbook, c As Chart, i As Long, n As Long
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    Util_FailThinSyncForTest = "delta-add:1"
    Core_ClearLastNotice
    Core_AddDelta "1:1", "1:3"
    Util_FailThinSyncForTest = ""
    For i = 1 To c.Shapes.Count
        If Left$(c.Shapes(i).Name, 14) = "tqq_ann_delta_" Then n = n + 1
    Next
    AssertEqual 0, n, "nada sobra"
    Teardown wb
End Sub

' Apagar a linha principal a mao: pernas e rotulo orfaos somem no proximo refresh; Remover setas apaga tudo.
Public Sub Test_Delta_Orphans_And_RemoveAll()
    Dim wb As Workbook, c As Chart
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "clustered")
    Core_AddDelta "1:1", "1:3"
    Core_AddDelta "2:1", "2:2"
    c.Shapes("tqq_ann_delta_1").Delete
    Ann_Refresh c
    AssertFalse HasShape(c, "tqq_ann_delta_1_lbl"): AssertFalse HasShape(c, "tqq_ann_delta_1_t1")
    AssertTrue HasShape(c, "tqq_ann_delta_2_lbl")
    c.Parent.Select
    Core_RemoveAnnotations
    AssertFalse HasShape(c, "tqq_ann_delta_2")
    Teardown wb
End Sub

' Cascata: texto pelo valor acumulado (Receita 100 -> Lucro 70) e perna da origem no topo da barra (Hi 100).
Public Sub Test_Delta_Waterfall_EndValues()
    Dim wb As Workbook, c As Chart
    Set wb = NewTempWorkbook()
    Set c = InsertOn(WaterfallFixture(wb), "waterfall")
    Core_AddDelta "0:1", "0:3"
    Relayout c
    AssertEqual Delta_Text(100, 70, "both", Style_NumberFormat(Registry_GetOption(c, "numfmt", "int"))), CStr(c.Shapes("tqq_ann_delta_1_lbl").TextFrame2.TextRange.Text)
    AssertAlmost YOf(c, 100) - 3, c.Shapes("tqq_ann_delta_1_t1").Top + c.Shapes("tqq_ann_delta_1_t1").Height, 1
    Teardown wb
End Sub

' Comparativa: delta entre etapas diferentes (A1 -> A3, fora do escopo das setas de variacao da 1.3.0).
Public Sub Test_Delta_WaterfallPair_BetweenSteps()
    Dim wb As Workbook, c As Chart, v1 As Double, v3 As Double, lo As Double, hi As Double
    Set wb = NewTempWorkbook()
    Set c = InsertOn(PairFixture(wb), "waterfallpair")
    Core_AddDelta "1:1", "1:3"
    Relayout c
    AssertTrue Delta_Extent(c, "1:1", v1, lo, hi) And Delta_Extent(c, "1:3", v3, lo, hi)
    AssertEqual Delta_Text(v1, v3, "both", Style_NumberFormat(Registry_GetOption(c, "numfmt", "int"))), CStr(c.Shapes("tqq_ann_delta_1_lbl").TextFrame2.TextRange.Text)
    AssertAlmost c.PlotArea.InsideLeft + c.PlotArea.InsideWidth * (1 - 0.5) / 8, c.Shapes("tqq_ann_delta_1_t1").Left, 1, "fatia A da etapa 1 (8 fatias)"
    AssertAlmost c.PlotArea.InsideLeft + c.PlotArea.InsideWidth * (7 - 0.5) / 8, c.Shapes("tqq_ann_delta_1_t2").Left, 1, "fatia A da etapa 3"
    Teardown wb
End Sub

' Comparativa com as setas de variacao do proprio par (tqq_pd) ligadas: a seta de delta passa ACIMA das elipses delas nas etapas do trecho
' (a borda conta o ponto de sala tqq_pdroom) e o rotulo do delta nao toca nenhuma elipse.
Public Sub Test_Delta_WaterfallPair_AboveVariationArrows()
    Dim wb As Workbook, c As Chart, pd As Series, i As Long, k As Long, dl As DataLabel, lbl As Shape, ln As Shape
    Set wb = NewTempWorkbook()
    Set c = InsertOn(PairFixture(wb), "waterfallpair")
    If Delta_FindSeriesByName(c, "tqq_pd") = 0 Then Wfp_SetArrows c, True
    AssertTrue Delta_FindSeriesByName(c, "tqq_pd") > 0, "setas de variacao ligadas"
    Core_AddDelta "1:1", "2:2"
    Relayout c
    Set ln = c.Shapes("tqq_ann_delta_1"): Set lbl = c.Shapes("tqq_ann_delta_1_lbl")
    Set pd = c.SeriesCollection(Delta_FindSeriesByName(c, "tqq_pd"))
    For i = 1 To 2
        k = 6 * (i - 1) + 3                     ' o ponto do meio do par (a elipse com o texto)
        If pd.Points(k).HasDataLabel Then
            Set dl = pd.Points(k).DataLabel
            AssertTrue ln.Top <= dl.Top - 0.5, "linha do delta acima da elipse da etapa " & i & " (" & ln.Top & " / " & dl.Top & ")"
            AssertTrue lbl.Top + lbl.Height <= dl.Top + 0.5 Or lbl.Left + lbl.Width <= dl.Left Or dl.Left + dl.Width <= lbl.Left, "rotulo do delta fora da elipse da etapa " & i
        End If
    Next
    AssertTrue lbl.Top >= 1.5, "rotulo dentro do grafico"
    Teardown wb
End Sub

' Empilhadas: a perna encosta na borda externa do SEGMENTO (Norte T1 vai de 10 a 15) e a linha passa acima da maior pilha do trecho (30).
Public Sub Test_Delta_Stacked_SegmentEdge()
    Dim wb As Workbook, c As Chart, lbl As Shape
    Set wb = NewTempWorkbook()
    Set c = InsertOn(Fixture(wb), "stacked")
    Core_AddDelta "2:1", "1:3"
    Relayout c
    Set lbl = c.Shapes("tqq_ann_delta_1_lbl")
    AssertAlmost YOf(c, 15) - 3, c.Shapes("tqq_ann_delta_1_t1").Top + c.Shapes("tqq_ann_delta_1_t1").Height, 1, "topo do segmento Norte T1 (10 + 5)"
    AssertAlmost YOf(c, 30) - 3, c.Shapes("tqq_ann_delta_1_t2").Top + c.Shapes("tqq_ann_delta_1_t2").Height, 1, "topo da pilha T3 (o Norte negativo desce)"
    AssertTrue lbl.Top + lbl.Height <= YOf(c, 30) - 8 + 0.5, "acima da maior pilha do trecho"
    Teardown wb
End Sub

' Combinacao com eixo secundario (a linha tem magnitude > 5x a das colunas): a ancora da linha e a borda do trecho usam o eixo DA SERIE.
Public Sub Test_Delta_Combo_SecondaryAxis()
    Dim wb As Workbook, ws As Worksheet, c As Chart, lbl As Shape, edge As Double, i As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array("Tri", "Sul", "Linha")
    ws.Range("A2:C2").Value = Array("T1", 10, 500)
    ws.Range("A3:C3").Value = Array("T2", 20, 600)
    ws.Range("A4:C4").Value = Array("T3", 30, 700)
    Set c = InsertOn(ws.Range("A1:C4"), "combo")
    If Registry_GetOption(c, "secondary", "0") <> "1" Then Combo_SetSecondary c, True
    AssertEqual CLng(xlSecondary), CLng(c.SeriesCollection(Delta_SeriesIndex(c, 2)).AxisGroup), "a linha no eixo secundario"
    Core_AddDelta "1:1", "2:3"
    Relayout c
    Set lbl = c.Shapes("tqq_ann_delta_1_lbl")
    AssertAlmost YOf(c, 10) - 3, c.Shapes("tqq_ann_delta_1_t1").Top + c.Shapes("tqq_ann_delta_1_t1").Height, 1, "coluna (eixo primario)"
    AssertAlmost YSecOf(c, 700) - 3, c.Shapes("tqq_ann_delta_1_t2").Top + c.Shapes("tqq_ann_delta_1_t2").Height, 1, "ponto da linha (eixo secundario)"
    edge = YOf(c, 30)
    For i = 5 To 7
        If YSecOf(c, i * 100) < edge Then edge = YSecOf(c, i * 100)
    Next
    AssertTrue lbl.Top + lbl.Height <= edge - 8 + 0.5, "acima das colunas e da linha do trecho"
    AssertTrue lbl.Top >= 1.5, "rotulo dentro do grafico"
    AssertEqual Delta_Text(10, 700, "both", Style_NumberFormat(Registry_GetOption(c, "numfmt", "int"))), CStr(lbl.TextFrame2.TextRange.Text)
    Teardown wb
End Sub

' Formato "mi" (milhoes: virgulas finais de escala): o texto segue o formato e a localidade.
Public Sub Test_Delta_Text_Millions()
    Dim f As String
    f = Style_NumberFormat("mi")
    AssertEqual "+" & Format$(1.5, "0.0") & " (+150%)", Delta_Text(1000000, 2500000, "both", f)
    AssertEqual "-" & Format$(1.5, "0.0"), Delta_Text(2500000, 1000000, "abs", f)
    AssertEqual "+2 (+50%)", Delta_Text(4000, 6000, "both", Style_NumberFormat("mil")), "milhares: uma virgula final"
End Sub

' ---------- PowerPoint (Parte A): o grafico colado leva as formas do delta ----------

Private Function ClipboardUsable(cht As Chart) As Boolean
    On Error Resume Next
    cht.CopyPicture xlScreen, xlPicture
    ClipboardUsable = (Err.Number = 0)
    Err.Clear
End Function

Private Sub WaitSeconds(s As Single)
    Dim t As Single
    t = Timer
    Do While Timer >= t And Timer - t < s
        DoEvents
    Loop
End Sub

' So cria o PowerPoint se nenhum processo POWERPNT existe, conferido ANTES de qualquer GetObject (nunca toca o PowerPoint do usuario).
Private Function StartScratchPowerPoint() As Object
    Dim app As Object, n As Long
    Ppt_ResetTestHooks
    PptSync_ResetTestHooks
    n = Ppt_ProcessCount()
    If n <> 0 Then Errors_Log "TestDelta: pulado (processo POWERPNT existente ou contagem indisponivel: " & n & ")": Exit Function
    On Error Resume Next
    Set app = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    If app Is Nothing Then Errors_Log "TestDelta: pulado (PowerPoint nao instalado)": Exit Function
    Set StartScratchPowerPoint = app
End Function

Private Function UserWindowAppeared(app As Object) As Boolean
    Dim n As Long, npv As Long
    n = -1: npv = -1
    On Error Resume Next
    n = app.Windows.Count
    npv = app.ProtectedViewWindows.Count
    On Error GoTo 0
    UserWindowAppeared = (n <> 0 Or npv <> 0)
    If UserWindowAppeared Then Errors_Log "TestDelta: resto do teste pulado (uma janela apareceu na instancia do teste: PowerPoint do usuario)"
End Function

Private Sub StopScratchPowerPoint(app As Object, own As Collection)
    Dim p As Variant, nw As Long, left_ As Long, t As Single, npv As Long
    On Error Resume Next
    Ppt_ResetTestHooks
    PptSync_ResetTestHooks
    If app Is Nothing Then Exit Sub
    For Each p In own
        nw = -1
        nw = p.Windows.Count
        If nw = 0 Then p.Saved = -1: p.Close
    Next
    Set p = Nothing
    Set own = New Collection
    left_ = -1: npv = -1
    left_ = app.Presentations.Count
    npv = app.ProtectedViewWindows.Count
    If left_ = 0 And npv = 0 Then
        app.Quit
        Set app = Nothing
        t = Timer
        Do While Ppt_ProcessCount() > 0
            If Timer < t Or Timer - t > 10 Then Exit Do
            WaitSeconds 0.25
        Loop
    Else
        Errors_Log "TestDelta: o PowerPoint do teste ficou aberto (" & left_ & " apresentacao(oes), " & npv & " janela(s) protegida(s))"
        Set app = Nothing
    End If
End Sub

' Regra de tres propria no grafico DO SLIDE (eixo de valores = Axes(2) do PowerPoint).
Private Function PptYOf(pc As Object, v As Double) As Double
    Dim ax As Object
    Set ax = pc.Axes(2)
    PptYOf = pc.PlotArea.InsideTop + pc.PlotArea.InsideHeight * (ax.MaximumScale - v) / (ax.MaximumScale - ax.MinimumScale)
End Function

' Colagem no PowerPoint (sincronizacao da Parte A): o grafico do slide tem as 4 formas do delta, o mesmo texto e a perna da origem 3 pt acima
' da barra Sul T1 medida pela escala DO SLIDE. Sem PowerPoint livre (processo do usuario aberto ou nao instalado): sai verde, motivo no log.
Public Sub Test_Delta_InPowerPoint_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, c As Chart, pc As Object, t1 As Object, sfx As Variant
    Dim own As New Collection, en As Long, ed As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0): own.Add pres
    Set slide = pres.Slides.Add(1, 12)
    Ppt_PinTargetForTest slide
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\delta-ppt.xlsx", xlOpenXMLWorkbook
    Set c = InsertOn(rng, "clustered")
    Core_AddDelta "1:1", "2:3"
    If Not ClipboardUsable(c) Then Errors_Log "TestDelta: colagem pulada (area de transferencia indisponivel)": GoTo Cleanup
    Ppt_SetAppForTest app
    If UserWindowAppeared(app) Then GoTo Cleanup
    Core_CopyToPowerPoint "chart"
    AssertEqual 1, slide.Shapes.Count, "grafico colado"
    Set pc = slide.Shapes(1).Chart
    For Each sfx In Array("", "_t1", "_t2", "_lbl")
        Set t1 = Nothing
        On Error Resume Next
        Set t1 = pc.Shapes("tqq_ann_delta_1" & sfx)
        On Error GoTo Cleanup
        AssertTrue Not t1 Is Nothing, "forma tqq_ann_delta_1" & sfx & " no slide"
    Next
    AssertEqual CStr(c.Shapes("tqq_ann_delta_1_lbl").TextFrame2.TextRange.Text), CStr(pc.Shapes("tqq_ann_delta_1_lbl").TextFrame2.TextRange.Text), "mesmo texto"
    Set t1 = pc.Shapes("tqq_ann_delta_1_t1")
    AssertAlmost PptYOf(pc, 10) - 3, CDbl(t1.Top + t1.Height), 1.5, "perna da origem 3 pt acima da barra Sul T1 no slide"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set t1 = Nothing: Set pc = Nothing: Set slide = Nothing: Set pres = Nothing
    StopScratchPowerPoint app, own
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestDelta", ed
End Sub
