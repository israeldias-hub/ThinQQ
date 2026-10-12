Attribute VB_Name = "modSamples"
Option Explicit
' Pasta de exemplos: uma planilha "Exemplos" com seis tabelas e seus graficos, um abaixo do outro
' (titulo na coluna A, a cada 21 linhas; a tabela ao lado do grafico, que nasce 2 colunas depois da tabela), e uma segunda
' planilha, "Exemplos 2", com os tipos e as decoracoes da Fase 2 (Samples_CreateSheet2), e uma terceira, "Exemplos 3", com os tipos e
' recursos da Fase 3 (Samples_CreateSheet3).

' Escreve uma matriz de linhas (Array de Arrays) a partir de topLeft; a primeira linha vira cabecalho em negrito.
Private Sub WriteTable(topLeft As Range, tbl As Variant)
    Dim i As Long, j As Long, r As Variant
    For i = LBound(tbl) To UBound(tbl)
        r = tbl(i)
        For j = LBound(r) To UBound(r)
            topLeft.Offset(i - LBound(tbl), j - LBound(r)).Value = r(j)
        Next
    Next
    r = tbl(LBound(tbl))
    topLeft.Resize(1, UBound(r) - LBound(r) + 1).Font.Bold = True
End Sub

Private Sub WriteTitle(ws As Worksheet, rowNo As Long, text As String)
    With ws.Cells(rowNo, 1)
        .Value = text
        .Font.Bold = True
        .Font.Size = 12
    End With
End Sub

' Tabela do exemplo 1 (Ano/Sul/Sudeste/Nordeste, 2022..2025), reaproveitada nas secoes 15 e 17 da segunda planilha.
Private Function YearTable() As Variant
    YearTable = Array(Array("Ano", "Sul", "Sudeste", "Nordeste"), Array(2022, 120, 210, 65), Array(2023, 131, 228, 72), _
                      Array(2024, 140, 251, 80), Array(2025, 158, 276, 91))
End Function

' Como depois de Core_Insert: o que o usuario faria em seguida (reposiciona rotulos, escala, anotacoes) ja vale na pasta de exemplos.
Private Sub Settle(co As ChartObject)
    Core_SafeRefresh co.Chart
End Sub

' Segunda planilha ("Exemplos 2", logo depois de "Exemplos"): os tipos e as decoracoes da Fase 2, uma secao a cada 21 linhas
' (titulo na coluna A, tabela abaixo dele, grafico 2 colunas a direita da tabela), numeradas de 7 a 18. Tudo pelos mesmos caminhos
' do usuario (Kinds_Insert, Deco_*, Pie_SetOutside, Ann_*), de modo que os exemplos tambem servem de teste de integracao.
' Deixa a planilha nova ativa; quem chama devolve a primeira. Propaga erros (Samples_Create fecha a pasta).
Private Function Samples_CreateSheet2(wb As Workbook) As Worksheet
    Dim ws As Worksheet, d As clsChartData, co As ChartObject, co2 As ChartObject, coll As Collection, skipped As Long
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets("Exemplos"))
    ws.Name = "Exemplos 2"
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B:E").ColumnWidth = 14

    WriteTitle ws, 1, "7) Barras empilhadas com totais e seta de diferença"
    WriteTable ws.Range("A2"), Array(Array("Região", "Loja", "Online"), Array("Sul", 10, 5), Array("Sudeste", 25, 10), Array("Norte", 40, 20))
    Set d = ChartData_FromRange(ws.Range("A2:C5"))
    Set co = Kinds_Insert(ws, d, "barstacked")
    Ann_AddDifference co.Chart, 1, 3
    Settle co

    WriteTitle ws, 22, "8) Cascata com várias séries (segmentos por série)"
    WriteTable ws.Range("A23"), Array(Array("Etapa", "Loja", "Online"), Array("Início", 100, 50), Array("Ajuste", -30, 10), _
                                      Array("Subtotal", "e", "e"), Array("Queda", -200, -20), Array("Fim", "e"))
    Set d = ChartData_FromRange(ws.Range("A23:C28"))
    Set co = Kinds_Insert(ws, d, "waterfall")
    Settle co

    WriteTitle ws, 43, "9) Cascata horizontal"
    WriteTable ws.Range("A44"), Array(Array("Etapa", "Valor"), Array("Receita", 1200), Array("Custos", -700), Array("Despesas", -250), _
                                      Array("EBITDA", "e"), Array("Impostos", -80), Array("Lucro", "e"))
    Set d = ChartData_FromRange(ws.Range("A44:B50"))
    Set co = Kinds_Insert(ws, d, "barwaterfall")
    Settle co

    WriteTitle ws, 64, "10) Mekko (%): largura = total da linha"
    WriteTable ws.Range("A65"), Array(Array("Região", "Eletro", "Móveis"), Array("Sul", 40, 10), Array("Sudeste", 20, 10), Array("Norte", 10, 10))
    Set d = ChartData_FromRange(ws.Range("A65:C68"))
    Set co = Kinds_Insert(ws, d, "mekko")
    Settle co

    WriteTitle ws, 85, "11) Mekko (unidades) com linha de valor"
    WriteTable ws.Range("A86"), Array(Array("Região", "Largura", "Eletro", "Móveis"), Array("Sul", 50, 40, 10), _
                                      Array("Sudeste", 30, 20, 10), Array("Norte", 20, 10, 10))
    Set d = ChartData_FromRange(ws.Range("A86:D89"))
    Set co = Kinds_Insert(ws, d, "mekkounits")
    Deco_SetValueLine co.Chart, True
    Settle co

    WriteTitle ws, 106, "12) Gantt com marco e colchete"
    WriteTable ws.Range("A107"), Array(Array("Atividade", "Início", "Fim"), Array("Planejamento", DateSerial(2026, 3, 2), DateSerial(2026, 3, 13)), _
                                       Array("Execução", DateSerial(2026, 3, 9), DateSerial(2026, 4, 17)), Array("Entrega", DateSerial(2026, 4, 20)))
    ws.Range("B108:C110").NumberFormat = "dd/mm/yyyy"
    Set d = ChartData_FromRange(ws.Range("A107:C110"))
    Set co = Kinds_Insert(ws, d, "gantt")
    Ann_AddBracket co.Chart, 1, 2
    Settle co

    WriteTitle ws, 127, "13) Rosca (rótulos sobre o anel)"
    WriteTable ws.Range("A128"), Array(Array("Produto", "Vendas"), Array("Eletro", 50), Array("Móveis", 30), Array("Tintas", 20))
    Set d = ChartData_FromRange(ws.Range("A128:B131"))
    Set co = Kinds_Insert(ws, d, "doughnut")
    Settle co

    WriteTitle ws, 148, "14) Pizza com rótulos fora (nomes e valores)"
    WriteTable ws.Range("A149"), Array(Array("Produto", "Vendas"), Array("Eletro", 50), Array("Móveis", 30), Array("Tintas", 20))
    Set d = ChartData_FromRange(ws.Range("A149:B152"))
    Set co = Kinds_Insert(ws, d, "pie")
    Pie_SetOutside co.Chart, True
    Settle co

    WriteTitle ws, 169, "15) Área empilhada"
    WriteTable ws.Range("A170"), YearTable()
    Set d = ChartData_FromRange(ws.Range("A170:D174"))
    Set co = Kinds_Insert(ws, d, "area")
    Settle co

    WriteTitle ws, 190, "16) Combinação: colunas + linha no eixo secundário"
    WriteTable ws.Range("A191"), Array(Array("Ano", "Vendas", "Custo", "Margem"), Array(2022, 100, 60, 0.42), Array(2023, 120, 70, 0.45), Array(2024, 130, 70, 0.47))
    ws.Range("D192:D194").NumberFormat = "0%"
    Set d = ChartData_FromRange(ws.Range("A191:D194"))
    Set co = Kinds_Insert(ws, d, "combo")
    Settle co

    WriteTitle ws, 211, "17) Colunas empilhadas com linha de valor (média)"
    WriteTable ws.Range("A212"), Array(Array("Ano", "Sul", "Sudeste", "Nordeste"), Array(2022, 118, 207, 65), Array(2023, 126, 212, 68), _
                                       Array(2024, 140, 251, 80), Array(2025, 158, 276, 91))
    Set d = ChartData_FromRange(ws.Range("A212:D216"))
    Set co = Kinds_Insert(ws, d, "stacked")
    Deco_SetValueLine co.Chart, True
    Settle co

    WriteTitle ws, 232, "18) Mesma escala: duas regiões lado a lado"
    WriteTable ws.Range("A233"), Array(Array("Ano (Sul)", "Loja", "Online"), Array(2023, 10, 40), Array(2024, 20, 50), Array(2025, 30, 60))
    WriteTable ws.Range("A239"), Array(Array("Ano (Sudeste)", "Loja", "Online"), Array(2023, 100, 400), Array(2024, 200, 500), Array(2025, 300, 600))
    Set d = ChartData_FromRange(ws.Range("A233:C236"))
    Set co = Kinds_Insert(ws, d, "clustered")
    co.Width = 300: co.Height = 200
    Set d = ChartData_FromRange(ws.Range("A239:C242"))
    Set co2 = Kinds_Insert(ws, d, "clustered")
    co2.Width = 300: co2.Height = 200
    co2.Left = co.Left + 320
    co2.Top = co.Top
    Settle co
    Settle co2
    ' tamanho e posicao definitivos ANTES da escala: a escala automatica do Excel depende da altura do grafico
    Set coll = New Collection
    coll.Add co.Chart
    coll.Add co2.Chart
    Deco_SameScale coll, skipped
    Set Samples_CreateSheet2 = ws
End Function

' Terceira planilha ("Exemplos 3"): os tipos e recursos da Fase 3, 8 secoes (19 a 26) a cada 21 linhas, pelos mesmos caminhos do usuario.
Private Function Samples_CreateSheet3(wb As Workbook) As Worksheet
    Dim ws As Worksheet, d As clsChartData, co As ChartObject
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets("Exemplos 2"))
    ws.Name = "Exemplos 3"
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B:F").ColumnWidth = 14

    WriteTitle ws, 1, "19) Borboleta: duas séries espelhadas"
    WriteTable ws.Range("A2"), Array(Array("Faixa", "Mulheres", "Homens"), Array("0-14", 120, 130), Array("15-29", 200, 210), Array("30-59", 310, 290), Array("60+", 90, 60))
    Set d = ChartData_FromRange(ws.Range("A2:C6"))
    Set co = Kinds_Insert(ws, d, "butterfly")
    Settle co

    WriteTitle ws, 22, "20) Dispersão por grupo, com tendência e quadrantes"
    WriteTable ws.Range("A23"), Array(Array("Produto", "Preço", "Margem", "Grupo"), Array("P1", 10, 0.2, "A"), Array("P2", 20, 0.35, "B"), _
                                      Array("P3", 30, 0.25, "A"), Array("P4", 40, 0.5, "B"), Array("P5", 25, 0.3, "A"), Array("P6", 35, 0.45, "B"))
    ws.Range("C24:C29").NumberFormat = "0%"
    Set d = ChartData_FromRange(ws.Range("A23:D29"))
    Set co = Kinds_Insert(ws, d, "scatter")
    Scatter_SetTrend co.Chart, True
    Scatter_SetQuad co.Chart, True
    Settle co

    WriteTitle ws, 43, "21) Bolhas: tamanho = volume"
    WriteTable ws.Range("A44"), Array(Array("Produto", "Preço", "Margem", "Volume", "Grupo"), Array("P1", 10, 0.2, 50, "A"), Array("P2", 20, 0.35, 120, "B"), _
                                      Array("P3", 30, 0.25, 80, "A"), Array("P4", 40, 0.5, 20, "B"))
    ws.Range("C45:C48").NumberFormat = "0%"
    Set d = ChartData_FromRange(ws.Range("A44:E48"))
    Set co = Kinds_Insert(ws, d, "bubble")
    Settle co

    WriteTitle ws, 64, "22) Colunas 100% com rótulos de soma consistente (34+33+33)"
    WriteTable ws.Range("A65"), Array(Array("Ano", "S1", "S2", "S3"), Array(2024, 1, 1, 1), Array(2025, 2, 3, 5))
    Set d = ChartData_FromRange(ws.Range("A65:D67"))
    Set co = Kinds_Insert(ws, d, "stacked100")
    Round_SetLabels co.Chart, True
    Settle co

    WriteTitle ws, 85, "23) Colunas empilhadas com corte de eixo em 50"
    WriteTable ws.Range("A86"), Array(Array("Cat", "S1", "S2"), Array("a", 5, 5), Array("b", 10, 10), Array("c", 20, 120))
    Set d = ChartData_FromRange(ws.Range("A86:C89"))
    Set co = Kinds_Insert(ws, d, "stacked")
    Break_Set co.Chart, 50
    Settle co

    WriteTitle ws, 106, "24) Barras empilhadas com corte de eixo"
    WriteTable ws.Range("A107"), Array(Array("Cat", "S1", "S2"), Array("a", 5, 5), Array("b", 10, 10), Array("c", 20, 120))
    Set d = ChartData_FromRange(ws.Range("A107:C110"))
    Set co = Kinds_Insert(ws, d, "barstacked")
    Break_Set co.Chart, 50
    Settle co

    WriteTitle ws, 127, "25) Área empilhada com soma consistente"
    WriteTable ws.Range("A128"), Array(Array("Ano", "Sul", "Norte", "Leste"), Array(2024, 10.4, 10.4, 10.4), Array(2025, 1, 1, 1), Array(2026, 2.5, 2.5, 2.5))
    Set d = ChartData_FromRange(ws.Range("A128:D131"))
    Set co = Kinds_Insert(ws, d, "area")
    Round_SetLabels co.Chart, True
    Settle co

    WriteTitle ws, 148, "26) Dispersão sem grupo, com quadrantes"
    WriteTable ws.Range("A149"), Array(Array("Produto", "Preço", "Margem"), Array("P1", 10, 0.2), Array("P2", 20, 0.35), Array("P3", 30, 0.25), _
                                       Array("P4", 40, 0.5), Array("P5", 25, 0.3), Array("P6", 35, 0.45))
    ws.Range("C150:C155").NumberFormat = "0%"
    Set d = ChartData_FromRange(ws.Range("A149:C155"))
    Set co = Kinds_Insert(ws, d, "scatter")
    Scatter_SetQuad co.Chart, True
    Settle co
    Set Samples_CreateSheet3 = ws
End Function

' Cria uma pasta nova com os exemplos e a devolve (ativa, A1 selecionada). Nao toca em outras pastas.
' Se algo falhar, a pasta nova e fechada sem salvar e o erro original e propagado.
Public Function Samples_Create() As Workbook
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, co As ChartObject
    Dim en As Long, es As String, ed As String, prevAlerts As Boolean
    On Error GoTo Fail
    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Exemplos"
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B:D").ColumnWidth = 14
    On Error Resume Next
    wb.Title = "ThinQQ - Exemplos"
    On Error GoTo Fail

    WriteTitle ws, 1, "1) Colunas empilhadas com totais, seta CAGR e seta de diferença (R$ mi)"
    WriteTable ws.Range("A2"), Array(Array("Ano", "Sul", "Sudeste", "Nordeste"), Array(2022, 120, 210, 65), _
                                     Array(2023, 131, 228, 72), Array(2024, 140, 251, 80), Array(2025, 158, 276, 91))
    Set d = ChartData_FromRange(ws.Range("A2:D6"))
    Set co = Stacked_Insert(ws, d, "stacked")
    Ann_AddCagr co.Chart, 1, 4
    Ann_AddDifference co.Chart, 1, 4

    WriteTitle ws, 22, "2) Cascata: escreva 'e' para uma barra de total calculada"
    WriteTable ws.Range("A23"), Array(Array("Etapa", "Valor"), Array("Receita", 1200), Array("Custos", -700), _
                                      Array("Despesas", -250), Array("EBITDA", "e"), Array("Impostos", -80), Array("Lucro", "e"))
    Set d = ChartData_FromRange(ws.Range("A23:B29"))
    Waterfall_Insert ws, d

    WriteTitle ws, 43, "3) Colunas agrupadas"
    WriteTable ws.Range("A44"), Array(Array("Trimestre", "Loja", "Online"), Array("T1", 420, 180), _
                                      Array("T2", 455, 230), Array("T3", 430, 290), Array("T4", 510, 360))
    Set d = ChartData_FromRange(ws.Range("A44:C48"))
    Stacked_Insert ws, d, "clustered"

    WriteTitle ws, 64, "4) Linha"
    WriteTable ws.Range("A65"), Array(Array("Mês", "Clientes ativos (mil)"), Array("jan", 48), Array("fev", 51), _
                                      Array("mar", 55), Array("abr", 54), Array("mai", 60), Array("jun", 66))
    Set d = ChartData_FromRange(ws.Range("A65:B71"))
    Stacked_Insert ws, d, "line"

    WriteTitle ws, 85, "5) Colunas 100% (participação de cada categoria de produto)"
    WriteTable ws.Range("A86"), Array(Array("Trimestre", "Eletro", "Móveis", "Tintas"), Array("T1", 50, 30, 20), _
                                      Array("T2", 45, 33, 22), Array("T3", 40, 35, 25), Array("T4", 38, 34, 28))
    Set d = ChartData_FromRange(ws.Range("A86:D90"))
    Stacked_Insert ws, d, "stacked100"

    WriteTitle ws, 106, "6) Cascata que cruza o zero, com subtotal 'e'"
    WriteTable ws.Range("A107"), Array(Array("Etapa", "Valor"), Array("Saldo inicial", 100), Array("Investimento", -180), _
                                       Array("Vendas", 50), Array("Subtotal", "e"), Array("Resgate", 90), Array("Saldo final", "e"))
    Set d = ChartData_FromRange(ws.Range("A107:B113"))
    Waterfall_Insert ws, d

    Samples_CreateSheet2 wb
    Samples_CreateSheet3 wb
    ws.Activate
    ws.Range("A1").Select
    Set Samples_Create = wb
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then
        prevAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        wb.Close SaveChanges:=False
        Application.DisplayAlerts = prevAlerts
    End If
    On Error GoTo 0
    Err.Raise en, es, ed
End Function
