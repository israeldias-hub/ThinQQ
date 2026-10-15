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
