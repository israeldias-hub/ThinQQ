Attribute VB_Name = "modSamples"
Option Explicit
' Pasta de exemplos: uma planilha "Exemplos" com seis tabelas e seus graficos, um abaixo do outro
' (titulo na coluna A, a cada 21 linhas; a tabela ao lado do grafico, que nasce 2 colunas depois da tabela).

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
