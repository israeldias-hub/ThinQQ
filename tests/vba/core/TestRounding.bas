Attribute VB_Name = "TestRounding"
Option Explicit

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

' Tabela 3 series x 2 categorias a partir de A1; devolve a faixa.
Private Function Fixture(ws As Worksheet, v As Variant) As Range
    ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Ano", "S1", "S2", "S3")
    ws.Range("A2:D2").Value = Array(2024, v(0), v(1), v(2))
    ws.Range("A3:D3").Value = Array(2025, v(3), v(4), v(5))
    Set Fixture = ws.Range("A1:D3")
End Function

Private Function Pct(n As Long, Optional dec As String = "") As String
    Pct = CStr(n) & dec & "%"
End Function

Public Sub Test_Round_Decimals_And_Formula()
    Dim wb As Workbook, ws As Worksheet, col As Variant, i As Long, f As String
    AssertEqual 0, Round_Decimals("int"): AssertEqual 1, Round_Decimals("dec1"): AssertEqual 2, Round_Decimals("dec2")
    AssertEqual 2, Round_Decimals("pct0"): AssertEqual 3, Round_Decimals("pct1"): AssertEqual -3, Round_Decimals("mil"): AssertEqual -5, Round_Decimals("mi")
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:C1").Value = Array(0.335, 0.335, 0.33): ws.Range("E1").Formula = "=SUM(A1:C1)": ws.Range("F1").Value = 2
    col = Array("A", "B", "C")
    For i = 0 To 2
        f = Round_Formula(col(i) & "1", "$A$1:$C$1", "$A$1:" & col(i) & "1", "$E$1", "$F$1")
        AssertEqual "=", Left$(f, 1): AssertTrue InStr(f, "SUMPRODUCT") > 0 And InStr(f, "COUNTIF") = 0, "so ROUND/SUMPRODUCT/POWER/IF"
        ws.Range(col(i) & "2").Formula = f
    Next
    AssertAlmost 0.33, ws.Range("A2").Value, 0.0000001: AssertAlmost 0.34, ws.Range("B2").Value, 0.0000001: AssertAlmost 0.33, ws.Range("C2").Value, 0.0000001
    ws.Range("A1:C1").Value = Array(1.5, 1.5, 1.5): ws.Range("F1").Value = 0
    AssertEqual 1, ws.Range("A2").Value: AssertEqual 2, ws.Range("B2").Value: AssertEqual 2, ws.Range("C2").Value
    ws.Range("A1:C1").Value = Array(33.3333, 33.3333, 33.3334)
    AssertEqual 33, ws.Range("A2").Value: AssertEqual 33, ws.Range("B2").Value: AssertEqual 34, ws.Range("C2").Value
    ws.Range("A1:C1").Value = Array(1234, 2345, 3456): ws.Range("F1").Value = -3
    AssertEqual 1000, ws.Range("A2").Value: AssertEqual 2000, ws.Range("B2").Value: AssertEqual 4000, ws.Range("C2").Value
    ws.Range("A1:C1").Value = Array(-1.5, 2.5, 2.5): ws.Range("F1").Value = 0
    AssertEqual -2, ws.Range("A2").Value: AssertEqual 3, ws.Range("B2").Value: AssertEqual 3, ws.Range("C2").Value
    CloseTemp wb
End Sub

' Review Focus 3: tres participacoes de um terco somam 100%.
Public Sub Test_Round_Stacked100_SharesSumTo100()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, blk As Range, body As Range, j As Long, s As Series, dec As String, sum As Double
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Set rng = Fixture(ws, Array(1, 1, 1, 2, 3, 5))
    rng.Select
    Core_Insert "stacked100"
    Set cht = Core_ActiveChart()
    AssertEqual Pct(33), cht.SeriesCollection(1).Points(1).DataLabel.Text, "antes: 33%"
    Core_SetToggle "roundsum", True
    AssertTrue Core_GetToggle("roundsum")
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertTrue Round_DCol(blk) > 0 And Round_RCol(blk, 3) > 0, "colunas RndD e Rnd3"
    AssertEqual 2, body.Cells(2, Round_DCol(blk)).Value, "pct0: 2 casas da fracao"
    AssertEqual "Pos", body.Cells(1, Stacked_PosCol(blk)).Value, "Pos continua achada pelo cabecalho"
    For j = 1 To 3
        Set s = cht.SeriesCollection(j)
        AssertTrue InStr(s.Points(1).DataLabel.Formula, CALC_SHEET) > 0, "rotulo " & j & " ligado ao bloco"
    Next
    AssertEqual Pct(34), cht.SeriesCollection(1).Points(1).DataLabel.Text, "empate: a primeira coluna recebe a unidade"
    AssertEqual Pct(33), cht.SeriesCollection(2).Points(1).DataLabel.Text
    AssertEqual Pct(33), cht.SeriesCollection(3).Points(1).DataLabel.Text
    sum = 0
    For j = 1 To 3: sum = sum + CDbl(body.Cells(2, Round_RCol(blk, j)).Value): Next
    AssertAlmost 1, sum, 0.0000001, "as participacoes arredondadas somam 100%"
    Core_SetOption "numfmt", "pct1"
    dec = Application.International(xlDecimalSeparator)
    AssertEqual 3, body.Cells(2, Round_DCol(blk)).Value, "pct1: 3 casas"
    AssertEqual Pct(33, dec & "4"), cht.SeriesCollection(1).Points(1).DataLabel.Text, "pct1: a primeira coluna recebe a unidade"
    AssertEqual Pct(33, dec & "3"), cht.SeriesCollection(3).Points(1).DataLabel.Text
    AssertEqual Pct(20, dec & "0"), cht.SeriesCollection(1).Points(2).DataLabel.Text, "linha 2: 2/10 exato"
    ws.Range("B3:D3").ClearContents
    cht.Refresh
    AssertEqual "", cht.SeriesCollection(1).Points(2).DataLabel.Text, "categoria vazia: sem 1% fantasma"
    ws.Range("B3:D3").Value = Array(2, 3, 5)
    Core_SetToggle "roundsum", False
    AssertFalse Core_GetToggle("roundsum")
    AssertEqual Pct(33, dec & "3"), cht.SeriesCollection(1).Points(1).DataLabel.Text, "desligado: participacao crua de novo (33,3, nao 33,4)"
    Teardown wb
End Sub

Public Sub Test_Round_Stacked_Values_SumToRoundedTotal()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, blk As Range, body As Range, cols As Long, j As Long, s As Series, src As Range
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Set rng = Fixture(ws, Array(1.5, 1.5, 1.5, 10, 20, 30))
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    cols = Registry_Block(cht).Columns.Count
    AssertEqual "5", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text, "total 4,5 mostra 5"
    AssertEqual "2", cht.SeriesCollection(1).Points(1).DataLabel.Text, "antes: cada 1,5 mostra 2 (2+2+2 = 6)"
    Core_SetToggle "roundsum", True
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual cols + 7, blk.Columns.Count, "RndD + 3 V + 3 Rnd"
    AssertEqual "V1", body.Cells(1, Round_VCol(blk, 1)).Value: AssertEqual "Rnd3", body.Cells(1, Round_RCol(blk, 3)).Value
    AssertEqual 1.5, body.Cells(2, Round_VCol(blk, 1)).Value, "V = valor saneado da origem"
    AssertEqual "1", cht.SeriesCollection(1).Points(1).DataLabel.Text
    AssertEqual "2", cht.SeriesCollection(2).Points(1).DataLabel.Text
    AssertEqual "2", cht.SeriesCollection(3).Points(1).DataLabel.Text, "1 + 2 + 2 = 5 = total arredondado"
    AssertEqual "5", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    AssertEqual "10", cht.SeriesCollection(1).Points(2).DataLabel.Text, "linha exata inalterada"
    ws.Range("B2").Value = "x"
    AssertEqual 0, body.Cells(2, Round_VCol(blk, 1)).Value, "texto vale 0 (sem #VALOR!)"
    AssertFalse IsError(body.Cells(2, Round_RCol(blk, 1)).Value)
    ws.Range("B2").Value = 1.5
    Set src = Block_SourceRange(cht)
    Core_SetOption "numfmt", "dec1"
    AssertEqual 1, body.Cells(2, Round_DCol(blk)).Value
    AssertEqual "1" & Application.International(xlDecimalSeparator) & "5", cht.SeriesCollection(1).Points(1).DataLabel.Text, "com 1 casa nada precisa de ajuste"
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(255, 255, 255), cht.SeriesCollection(1).DataLabels.Font.Color, "fonte por contraste continua"
    Teardown wb
End Sub

Public Sub Test_Round_Area_Toggle_And_Rebuild()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, blk As Range, lbl As Series, before As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Set rng = Fixture(ws, Array(10.4, 10.4, 10.4, 1, 1, 1))
    rng.Select
    Core_Insert "area"
    Set cht = Core_ActiveChart()
    AssertEqual "31", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    Core_SetToggle "roundsum", True
    Set blk = Registry_Block(cht)
    AssertEqual 3, Area_SeriesCount(blk), "contagem por cabecalho sobrevive a extensao"
    Set lbl = Chart_FindSeries(cht, "tqq_lbl_3")
    AssertTrue InStr(lbl.Points(1).DataLabel.Formula, Block_Body(blk).Cells(2, Round_RCol(blk, 3)).Address(True, True)) > 0, "rotulo de segmento ligado a Rnd3"
    AssertEqual "11", Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text, "empate: a primeira coluna recebe a unidade"
    AssertEqual "10", lbl.Points(1).DataLabel.Text, "11 + 10 + 10 = 31"
    AssertTrue Core_RebuildFromRange(cht, rng), "rebuild"
    Set cht = Core_ActiveChart()
    AssertTrue Core_GetToggle("roundsum"), "opcao mantida"
    Set blk = Registry_Block(cht)
    AssertTrue Round_RCol(blk, 3) > 0, "colunas recriadas no bloco novo"
    AssertEqual "11", Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text, "religado apos o rebuild"
    Core_SetToggle "roundsum", False
    AssertTrue InStr(Chart_FindSeries(cht, "tqq_lbl_3").Points(1).DataLabel.Formula, Block_Body(blk).Cells(2, Area_ValCol(blk, 3)).Address(True, True)) > 0, "desligado: Val3"
    before = blk.Columns.Count
    Core_SetToggle "roundsum", True
    AssertEqual before, Registry_Block(cht).Columns.Count, "religar nao estende de novo"
    Teardown wb
End Sub

Public Sub Test_Round_Refusals_And_LabelsOff()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, en As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Set rng = Fixture(ws, Array(1, 2, 3, 4, 5, 6))
    rng.Select
    Core_Insert "clustered"
    Set cht = Core_ActiveChart()
    Core_ClearLastNotice
    Core_SetToggle "roundsum", True
    AssertEqual "'Rótulos com soma consistente' não se aplica a este tipo de gráfico.", Core_LastNotice()
    AssertFalse Core_GetToggle("roundsum")
    AssertEqual 0, Round_DCol(Registry_Block(cht)), "nada gravado no bloco"
    On Error Resume Next
    Round_SetLabels cht, True
    en = Err.Number
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    rng.Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetToggle "labels", False
    Core_SetToggle "roundsum", True
    AssertFalse cht.SeriesCollection(1).HasDataLabels, "soma consistente com rotulos desligados nao liga rotulos"
    Core_SetToggle "labels", True
    AssertTrue InStr(cht.SeriesCollection(1).Points(1).DataLabel.Formula, CALC_SHEET) > 0, "religar os rotulos liga a Rnd"
    Teardown wb
End Sub

Public Sub Test_Round_WriteFormulas_Range_And_Validate()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:A3").Value = Application.Transpose(Array(0.335, 0.335, 0.33))
    AssertEqual "", Round_ValidateRange(ws.Range("A1:A3"))
    AssertEqual "Selecione uma única coluna ou linha só com números (pelo menos 2).", Round_ValidateRange(ws.Range("A1"))
    AssertEqual "Selecione uma única coluna ou linha só com números (pelo menos 2).", Round_ValidateRange(ws.Range("A1:B3"))
    ws.Range("B2").Value = "x"
    AssertEqual "Selecione uma única coluna ou linha só com números (pelo menos 2).", Round_ValidateRange(ws.Range("B1:B3"))
    Round_WriteFormulas ws.Range("A1:A3"), ws.Range("C1"), 2
    AssertAlmost 0.33, ws.Range("C1").Value, 0.0000001: AssertAlmost 0.34, ws.Range("C2").Value, 0.0000001: AssertAlmost 0.33, ws.Range("C3").Value, 0.0000001
    AssertTrue InStr(ws.Range("C1").Formula, "$A$1:$A$3") > 0 And InStr(ws.Range("C1").Formula, CALC_SHEET) = 0, "so a origem"
    ws.Range("E1:G1").Value = Array(1234, 2345, 3456)
    Round_WriteFormulas ws.Range("E1:G1"), ws.Range("E3"), -3
    AssertEqual 1000, ws.Range("E3").Value: AssertEqual 2000, ws.Range("F3").Value: AssertEqual 4000, ws.Range("G3").Value
    ws.Range("A1").Value = 0.5
    AssertAlmost 0.5, ws.Range("C1").Value, 0.0000001, "vivo: a origem mudou (0,5 + 0,335 + 0,33 = 1,165 -> 1,17: 0,5 + 0,34 + 0,33)"
    AssertAlmost 1.17, ws.Range("C1").Value + ws.Range("C2").Value + ws.Range("C3").Value, 0.0000001
    CloseTemp wb
End Sub

Public Sub Test_Core_RoundRange_TestBuild_ValidatesWithoutInputBox()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1").Select
    Core_ClearLastNotice
    Core_RoundRange
    AssertEqual "Selecione uma única coluna ou linha só com números (pelo menos 2).", Core_LastNotice()
    ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))
    ws.Range("A1:A3").Select
    Core_ClearLastNotice
    Core_RoundRange                                             ' build Test: sai antes do InputBox, sem aviso
    AssertEqual "", Core_LastNotice()
    AssertEqual 0, Application.WorksheetFunction.CountA(ws.Range("B1:Z10")), "sem respostas: nada gravado"
    ' com as respostas do gancho (as dos dois InputBox): grava a partir de C1 e avisa o endereco
    Core_TestSetRoundRangeAnswers 0, ws.Range("C1")
    Core_ClearLastNotice
    Core_RoundRange
    AssertEqual "Fórmulas de arredondamento com soma consistente gravadas em C1:C3.", Core_LastNotice()
    AssertEqual 1, ws.Range("C1").Value: AssertEqual 2, ws.Range("C2").Value: AssertEqual 3, ws.Range("C3").Value
    AssertTrue InStr(ws.Range("C2").Formula, "SUMPRODUCT") > 0, "formula viva"
    ' destino preenchido: pergunta (o build Test responde Nao) e nada muda
    ws.Range("E2").Value = "manter"
    Core_TestSetRoundRangeAnswers 0, ws.Range("E1")
    Core_ClearLastNotice
    Core_RoundRange
    AssertEqual "As células E1:E3 já têm conteúdo. Substituir? Esta ação não pode ser desfeita.", Core_LastConfirm()
    AssertEqual "", Core_LastNotice(), "Nao: sem aviso de sucesso"
    AssertEqual "manter", ws.Range("E2").Value
    AssertTrue IsEmpty(ws.Range("E1").Value), "nada gravado"
    ' casas decimais invalidas e destino sobre a origem: avisos, nada gravado; o gancho e de uso unico
    Core_TestSetRoundRangeAnswers 1.5, ws.Range("G1")
    Core_RoundRange
    AssertEqual "Informe um número inteiro de casas decimais (de -15 a 15).", Core_LastNotice()
    AssertTrue IsEmpty(ws.Range("G1").Value)
    Core_TestSetRoundRangeAnswers 0, ws.Range("A2")
    Core_RoundRange
    AssertEqual "O resultado não pode ficar sobre a faixa de origem.", Core_LastNotice()
    AssertEqual 2, ws.Range("A2").Value
    Core_ClearLastNotice
    Core_RoundRange
    AssertEqual "", Core_LastNotice(), "gancho de uso unico: sem respostas, sai"
    CloseTemp wb
End Sub

' Falha depois de estender o bloco (gancho logo depois de gravar as formulas, antes do cabecalho RndD): a extensao e desfeita (o nome volta a
' faixa anterior, nada fica a direita), a opcao nao e gravada, os rotulos ficam como estavam e o gancho nao fica armado. Sem o gancho, liga.
Public Sub Test_Round_EnsureColumns_FailMidway_RollsBack()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, blk As Range, cols As Long, addr As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Fixture(ws, Array(1.5, 1.5, 1.5, 10, 20, 30)).Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Set blk = Registry_Block(cht): cols = blk.Columns.Count: addr = blk.Address
    Core_ClearLastNotice
    Util_FailThinSyncForTest = "round-ensure:write"
    Core_SetToggle "roundsum", True
    AssertEqual "", Util_FailThinSyncForTest, "o gancho dispara uma vez"
    Util_FailThinSyncForTest = ""
    AssertTrue InStr(Core_LastNotice(), "Falha simulada") > 0, "o erro chega ao usuario: " & Core_LastNotice()
    AssertFalse Core_GetToggle("roundsum"), "opcao nao gravada"
    Set blk = Registry_Block(cht)
    AssertEqual addr, blk.Address, "o nome volta a faixa anterior"
    AssertEqual 0, Application.WorksheetFunction.CountA(blk.Offset(0, cols).Resize(blk.Rows.Count, 7)), "nada a direita do bloco"
    AssertEqual 0, Round_DCol(blk)
    AssertEqual "2", cht.SeriesCollection(1).Points(1).DataLabel.Text, "rotulo nativo como antes"
    Core_SetToggle "roundsum", True
    AssertTrue Core_GetToggle("roundsum")
    AssertEqual cols + 7, Registry_Block(cht).Columns.Count, "estende uma vez so, a partir da faixa original"
    AssertEqual "1", cht.SeriesCollection(1).Points(1).DataLabel.Text
    Teardown wb
End Sub

' Falha no meio de religar os rotulos (depois de desligar os nativos, antes de ligar as Rnd): a opcao volta a "0" e os rotulos nativos sao
' refeitos (nenhuma serie fica sem rotulo); as colunas criadas ficam (religar nao estende de novo).
Public Sub Test_Round_Relink_FailMidway_RestoresLabels()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, j As Long, cols As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Fixture(ws, Array(1.5, 1.5, 1.5, 10, 20, 30)).Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_ClearLastNotice
    Util_FailThinSyncForTest = "round-relink:cleared"
    Core_SetToggle "roundsum", True
    AssertEqual "", Util_FailThinSyncForTest, "o gancho dispara uma vez (a limpeza passa pelo mesmo ponto)"
    Util_FailThinSyncForTest = ""
    AssertTrue InStr(Core_LastNotice(), "Falha simulada") > 0, "o erro chega ao usuario: " & Core_LastNotice()
    AssertEqual "0", Registry_GetOption(cht, "roundsum"), "opcao anterior"
    For j = 1 To 3
        AssertTrue cht.SeriesCollection(j).HasDataLabels, "serie " & j & " com rotulos"
        AssertTrue InStr(cht.SeriesCollection(j).Points(1).DataLabel.Formula, CALC_SHEET) = 0, "serie " & j & ": rotulo nativo"
    Next
    AssertEqual "2", cht.SeriesCollection(1).Points(1).DataLabel.Text
    cols = Registry_Block(cht).Columns.Count
    Core_SetToggle "roundsum", True
    AssertEqual cols, Registry_Block(cht).Columns.Count, "colunas reaproveitadas"
    AssertEqual "1", cht.SeriesCollection(1).Points(1).DataLabel.Text
    Teardown wb
End Sub

' Empilhadas: os valores de origem vem da faixa viva da formula do Total, nao do endereco gravado na meta: a planilha renomeada e uma linha
' inserida acima da tabela depois de inserir o grafico nao quebram nem desalinham as colunas V_j.
Public Sub Test_Round_Stacked_SourceRenamedAndShifted()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, blk As Range, body As Range
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Fixture(ws, Array(1.5, 1.5, 1.5, 10, 20, 30)).Select
    Core_Insert "barstacked"
    Set cht = Core_ActiveChart()
    ws.Name = "Dados novos"
    ws.Rows(1).Insert
    Core_ClearLastNotice
    Core_SetToggle "roundsum", True
    AssertEqual "", Core_LastNotice(), "sem aviso"
    AssertTrue Core_GetToggle("roundsum")
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    AssertEqual 1.5, body.Cells(2, Round_VCol(blk, 1)).Value, "V1 da linha 2024"
    AssertEqual 30, body.Cells(3, Round_VCol(blk, 3)).Value, "V3 da linha 2025"
    AssertEqual "1", cht.SeriesCollection(1).Points(1).DataLabel.Text
    AssertEqual "2", cht.SeriesCollection(3).Points(1).DataLabel.Text
    ws.Range("B3").Value = 2.4
    AssertEqual "2", cht.SeriesCollection(1).Points(1).DataLabel.Text, "vivo: 2,4 + 1,5 + 1,5 = 5,4 -> 5 = 2 + 2 + 1"
    AssertEqual "1", cht.SeriesCollection(2).Points(1).DataLabel.Text, "empate de restos -0,5: a primeira empatada perde a unidade"
    AssertEqual "2", cht.SeriesCollection(3).Points(1).DataLabel.Text
    Teardown wb
End Sub

' Area 100% (participacoes, RndD = 2 mesmo com numfmt int/dec1) e barras empilhadas em milhares (casas negativas).
Public Sub Test_Round_Area100_And_Thousands()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, blk As Range
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Fixture(ws, Array(1, 1, 1, 2, 3, 5)).Select
    Core_Insert "area100"
    Set cht = Core_ActiveChart()
    Core_SetToggle "roundsum", True
    Set blk = Registry_Block(cht)
    AssertEqual 2, Block_Body(blk).Cells(2, Round_DCol(blk)).Value
    AssertEqual Pct(34), Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text
    AssertEqual Pct(33), Chart_FindSeries(cht, "tqq_lbl_3").Points(1).DataLabel.Text
    Core_SetOption "numfmt", "dec1"
    AssertEqual 2, Block_Body(blk).Cells(2, Round_DCol(blk)).Value, "100%: os segmentos continuam em pct0"
    AssertEqual Pct(34), Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    ws.Range("B2:D2").Value = Array(1234, 2345, 3456)
    ws.Range("A1:D3").Select
    Core_Insert "barstacked"
    Set cht = Core_ActiveChart()
    Core_SetOption "numfmt", "mil"
    Core_SetToggle "roundsum", True
    Set blk = Registry_Block(cht)
    AssertEqual -3, Block_Body(blk).Cells(2, Round_DCol(blk)).Value
    AssertEqual "1", cht.SeriesCollection(1).Points(1).DataLabel.Text
    AssertEqual "2", cht.SeriesCollection(2).Points(1).DataLabel.Text
    AssertEqual "4", cht.SeriesCollection(3).Points(1).DataLabel.Text, "1 + 2 + 4 = 7 = total 7.035 em milhares"
    AssertEqual "7", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    Teardown wb
End Sub

Public Sub Test_Round_ValidateDest()
    Dim wb As Workbook, wb2 As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))
    AssertEqual "", Round_ValidateDest(ws.Range("A1:A3"), ws.Range("B1"))
    AssertEqual "", Round_ValidateDest(ws.Range("A1:A3"), ws.Range("A4"))
    AssertEqual "O resultado não pode ficar sobre a faixa de origem.", Round_ValidateDest(ws.Range("A1:A3"), ws.Range("A3"))
    AssertEqual "O resultado não cabe na planilha a partir dessa célula.", Round_ValidateDest(ws.Range("A1:A3"), ws.Cells(ws.Rows.Count - 1, 2))
    Set wb2 = NewTempWorkbook()
    AssertEqual "Escolha a célula de destino na mesma pasta de trabalho da faixa de origem.", Round_ValidateDest(ws.Range("A1:A3"), wb2.Worksheets(1).Range("A1"))
    CloseTemp wb2
    CloseTemp wb
End Sub

' Arredondar intervalo acima do limite (2.000 celulas): recusa com aviso, nada gravado (pelo comando e chamando a gravacao direto).
Public Sub Test_Round_RangeAboveCap_RefusesWithoutWriting()
    Dim wb As Workbook, ws As Worksheet, en As Long, ed As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:A2001").Value = 1
    AssertEqual Round_TooBigText(2001), Round_ValidateRange(ws.Range("A1:A2001"))
    AssertTrue InStr(Round_TooBigText(2001), "no máximo") > 0
    AssertEqual "", Round_ValidateRange(ws.Range("A1:A2000")), "2.000 celulas ainda aceitas"
    ws.Range("A1:A2001").Select
    Core_ClearLastNotice
    Core_RoundRange
    AssertEqual Round_TooBigText(2001), Core_LastNotice()
    On Error Resume Next
    Round_WriteFormulas ws.Range("A1:A2001"), ws.Range("C1"), 0
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    AssertEqual Round_TooBigText(2001), ed
    AssertEqual 0, Application.WorksheetFunction.CountA(ws.Range("C:C")), "nada gravado"
    CloseTemp wb
End Sub

' A gravacao devolve o modo de calculo e a atualizacao de tela anteriores (exatos) depois de gravar com sucesso, e o resultado sai
' calculado mesmo em calculo manual.
Public Sub Test_Round_WriteFormulas_RestoresCalcAndScreen()
    Dim wb As Workbook, ws As Worksheet, prevCalc As Long, prevSU As Boolean
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    prevCalc = Application.Calculation: prevSU = Application.ScreenUpdating
    ws.Range("A1:A3").Value = Application.Transpose(Array(0.335, 0.335, 0.33))
    Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
    Round_WriteFormulas ws.Range("A1:A3"), ws.Range("C1"), 2
    AssertEqual xlCalculationAutomatic, Application.Calculation, "automatico de volta"
    AssertTrue Application.ScreenUpdating, "tela de volta"
    AssertAlmost 0.34, ws.Range("C2").Value, 0.0000001
    Application.Calculation = xlCalculationManual: Application.ScreenUpdating = False
    Round_WriteFormulas ws.Range("A1:A3"), ws.Range("E1"), 2
    AssertEqual xlCalculationManual, Application.Calculation, "manual de volta"
    AssertFalse Application.ScreenUpdating, "tela desligada de volta"
    AssertAlmost 0.33, ws.Range("E1").Value, 0.0000001, "calculado em manual"
    AssertAlmost 0.34, ws.Range("E2").Value, 0.0000001, "calculado em manual"
    Application.Calculation = prevCalc: Application.ScreenUpdating = prevSU
    CloseTemp wb
End Sub

' Falha no meio da gravacao (gancho de disparo unico depois das formulas): o conteudo e os formatos anteriores do destino voltam, o modo de
' calculo e a tela voltam, o erro chega ao chamador e o gancho nao fica armado.
Public Sub Test_Round_WriteFormulas_FailMidway_RestoresAll()
    Dim wb As Workbook, ws As Worksheet, prevCalc As Long, prevSU As Boolean, en As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    prevCalc = Application.Calculation: prevSU = Application.ScreenUpdating
    ws.Range("A1:A3").Value = Application.Transpose(Array(0.335, 0.335, 0.33))
    ws.Range("C1").Value = "a": ws.Range("C2").Formula = "=1+1": ws.Range("C3").Value = 0.5: ws.Range("C3").NumberFormat = "0.0%"
    Application.Calculation = xlCalculationAutomatic: Application.ScreenUpdating = True
    Util_FailThinSyncForTest = "round-write:formulas"
    On Error Resume Next
    Round_WriteFormulas ws.Range("A1:A3"), ws.Range("C1"), 2
    en = Err.Number
    On Error GoTo 0
    AssertEqual "", Util_FailThinSyncForTest, "o gancho dispara uma vez"
    Util_FailThinSyncForTest = ""
    AssertEqual vbObjectError + 9995, en, "o erro chega ao chamador"
    AssertEqual "a", ws.Range("C1").Value
    AssertEqual "=1+1", ws.Range("C2").Formula
    AssertEqual 0.5, ws.Range("C3").Value
    AssertEqual "0.0%", ws.Range("C3").NumberFormat
    AssertEqual xlCalculationAutomatic, Application.Calculation
    AssertTrue Application.ScreenUpdating
    Application.Calculation = xlCalculationManual
    Util_FailThinSyncForTest = "round-write:formulas"
    On Error Resume Next
    Round_WriteFormulas ws.Range("A1:A3"), ws.Range("C1"), 2
    On Error GoTo 0
    Util_FailThinSyncForTest = ""
    AssertEqual xlCalculationManual, Application.Calculation, "manual de volta tambem no erro"
    AssertEqual "a", ws.Range("C1").Value
    Application.Calculation = prevCalc: Application.ScreenUpdating = prevSU
    CloseTemp wb
End Sub

' Destino com celula mesclada ou protegida: recusa limpa (7002 com a mensagem), nada gravado.
Public Sub Test_Round_WriteFormulas_MergedOrProtectedDest_Refuses()
    Dim wb As Workbook, ws As Worksheet, en As Long, ed As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:A3").Value = Application.Transpose(Array(1, 2, 3))
    ws.Range("C2:D2").Merge
    AssertEqual "O destino tem células mescladas; escolha uma área sem mesclagem.", Round_ValidateDest(ws.Range("A1:A3"), ws.Range("C1"))
    On Error Resume Next
    Round_WriteFormulas ws.Range("A1:A3"), ws.Range("C1"), 0
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    AssertEqual vbObjectError + 7002, en
    AssertEqual "O destino tem células mescladas; escolha uma área sem mesclagem.", ed
    AssertEqual 0, Application.WorksheetFunction.CountA(ws.Range("C1:C3")), "nada gravado"
    ws.Range("F1:F3").Locked = False
    ws.Protect
    AssertEqual "", Round_ValidateDest(ws.Range("A1:A3"), ws.Range("F1")), "celulas desbloqueadas servem"
    AssertEqual "As células de destino estão protegidas.", Round_ValidateDest(ws.Range("A1:A3"), ws.Range("G1"))
    AssertEqual "As células de destino estão protegidas.", Round_ValidateDest(ws.Range("A1:A3"), ws.Range("F2")), "F4 bloqueada"
    ws.Unprotect
    CloseTemp wb
End Sub

' Pergunta Sim/Nao do core (Core_Confirm): no build de teste so registra e responde Nao.
Public Sub Test_Core_Confirm_TestBuild_AnswersNo()
    Core_ClearLastNotice
    AssertEqual "", Core_LastConfirm()
    AssertFalse Core_Confirm("Substituir?"), "build Test: Nao"
    AssertEqual "Substituir?", Core_LastConfirm()
    AssertEqual "", Core_LastNotice(), "nao e um aviso"
End Sub

' Empate de restos com 1 casa: (2,25 - 2,3) / 0,1 e (2,05 - 2,1) / 0,1 diferem no ultimo bit (-0,499999999999998 contra -0,500000000000003,
' medido); os restos sao arredondados a 9 casas antes de comparar, entao vale a regra do empate (a PRIMEIRA coluna perde a unidade com K = -1).
' 0,25 / 0,35 (restos -0,5 exatos no Excel) seguem a mesma regra: 0,2 / 0,4.
Public Sub Test_Round_TieWithFloatNoise_FirstColumnLoses()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, dec As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Fixture(ws, Array(2.25, 2.05, 0, 0.25, 0.35, 0)).Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetOption "numfmt", "dec1"
    Core_SetToggle "roundsum", True
    dec = Application.International(xlDecimalSeparator)
    AssertEqual "2" & dec & "2", cht.SeriesCollection(1).Points(1).DataLabel.Text, "2,25 / 2,05 (total 4,3): a primeira perde a unidade"
    AssertEqual "2" & dec & "1", cht.SeriesCollection(2).Points(1).DataLabel.Text
    AssertEqual "4" & dec & "3", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    AssertEqual "0" & dec & "2", cht.SeriesCollection(1).Points(2).DataLabel.Text, "0,25 / 0,35 (total 0,6): 0,2 / 0,4"
    AssertEqual "0" & dec & "4", cht.SeriesCollection(2).Points(2).DataLabel.Text
    Teardown wb
End Sub

' Empilhadas: S vem das colunas que o Total soma (a faixa da formula do Total), nao das series do grafico. Uma serie apagada a mao antes de
' ligar: recusa com aviso (rotulo j -> Rnd_j pela ordem nao bate), opcao nao gravada. Apagada depois de ligar: religar os rotulos volta aos
' rotulos sem a opcao (nenhum rotulo ligado a coluna errada).
Public Sub Test_Round_Stacked_SeriesDeletedByHand()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, blk As Range, cols As Long, j As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Fixture(ws, Array(1.5, 1.5, 1.5, 10, 20, 30)).Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    cols = Registry_Block(cht).Columns.Count
    cht.SeriesCollection(2).Delete
    Core_ClearLastNotice
    Core_SetToggle "roundsum", True
    AssertEqual "As séries do gráfico não correspondem mais às colunas da tabela de origem; use 'Alterar dados' para reconstruí-lo.", Core_LastNotice()
    AssertFalse Core_GetToggle("roundsum")
    Set blk = Registry_Block(cht)
    AssertEqual cols + 7, blk.Columns.Count, "3 colunas V/Rnd: as que o Total soma"
    AssertTrue InStr(Block_Body(blk).Cells(2, Round_VCol(blk, 3)).Formula, ",1,3)") > 0, "V3 = 3a coluna da faixa do Total"
    For j = 1 To 2
        AssertTrue InStr(cht.SeriesCollection(j).Points(1).DataLabel.Formula, CALC_SHEET) = 0, "serie " & j & ": rotulo nativo"
    Next
    ws.ChartObjects(1).Delete
    CalcSheet_Cleanup wb
    Fixture(ws, Array(1.5, 1.5, 1.5, 10, 20, 30)).Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetToggle "roundsum", True
    AssertEqual "1", cht.SeriesCollection(1).Points(1).DataLabel.Text
    cht.SeriesCollection(1).Delete
    Core_SetToggle "labels", False
    Core_SetToggle "labels", True
    For j = 1 To 2
        AssertTrue InStr(cht.SeriesCollection(j).Points(1).DataLabel.Formula, CALC_SHEET) = 0, "serie " & j & ": rotulo nativo depois de apagar uma serie"
    Next
    AssertEqual "2", cht.SeriesCollection(1).Points(1).DataLabel.Text, "1,5 nativo"
    Teardown wb
End Sub

' Origem em percentual: a saida continua em percentual, com as casas do arredondamento (2 casas = 0%, 3 = 0,0%). Texto digitado depois
' na origem: #VALOR! (documentado), e volta ao numero quando a celula volta a ser numero.
Public Sub Test_Round_WriteFormulas_PercentSource_KeepsPercent()
    Dim wb As Workbook, ws As Worksheet, dec As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    dec = Application.International(xlDecimalSeparator)
    ws.Range("A1:A3").Value = Application.Transpose(Array(0.335, 0.335, 0.33))
    ws.Range("A1:A3").NumberFormat = "0.0%"
    Round_WriteFormulas ws.Range("A1:A3"), ws.Range("C1"), 2
    AssertEqual "33%", ws.Range("C1").Text: AssertEqual "34%", ws.Range("C2").Text: AssertEqual "33%", ws.Range("C3").Text
    ws.Range("A1:A3").Value = Application.Transpose(Array(1 / 3, 1 / 3, 1 / 3))
    Round_WriteFormulas ws.Range("A1:A3"), ws.Range("E1"), 3
    AssertEqual "33" & dec & "4%", ws.Range("E1").Text: AssertEqual "33" & dec & "3%", ws.Range("E3").Text
    ws.Range("A2").Value = "x"
    AssertTrue IsError(ws.Range("E1").Value), "texto na origem: #VALOR!"
    ws.Range("A2").Value = 1 / 3
    AssertAlmost 0.334, ws.Range("E1").Value, 0.0000001, "volta"
    CloseTemp wb
End Sub

' "Alterar dados" com a soma consistente ligada: as colunas Rnd_j sao criadas antes da passada dos rotulos, que ja liga a elas; a passada
' de roundsum nao refaz os rotulos (cada religacao custa Labels_LinkPoints ponto a ponto: ~17 s em 300 pontos de um 100%).
Public Sub Test_Round_Rebuild_LinksLabelsOnce()
    Dim wb As Workbook, ws As Worksheet, rng As Range, cht As Chart, k As Variant
    For Each k In Array("stacked100", "area")
        Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
        Set rng = Fixture(ws, Array(1, 1, 1, 2, 3, 5))
        rng.Select
        Core_Insert CStr(k)
        Set cht = Core_ActiveChart()
        Core_SetToggle "roundsum", True
        Round_ResetRelinkCountForTest
        AssertTrue Core_RebuildFromRange(cht, rng), k & ": rebuild"
        Set cht = Core_ActiveChart()
        AssertEqual 0, Round_RelinkCountForTest(), k & ": a passada de roundsum nao religa"
        AssertTrue Core_GetToggle("roundsum"), CStr(k)
        If k = "area" Then
            AssertTrue InStr(Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Formula, Block_Body(Registry_Block(cht)).Cells(2, Round_RCol(Registry_Block(cht), 1)).Address(True, True)) > 0, k & ": ligado a Rnd1"
            AssertEqual "1", Chart_FindSeries(cht, "tqq_lbl_1").Points(1).DataLabel.Text, k & ": 1 + 1 + 1 inalterados"
        Else
            AssertTrue InStr(cht.SeriesCollection(1).Points(1).DataLabel.Formula, Block_Body(Registry_Block(cht)).Cells(2, Round_RCol(Registry_Block(cht), 1)).Address(True, True)) > 0, k & ": ligado a Rnd1"
            AssertEqual Pct(34), cht.SeriesCollection(1).Points(1).DataLabel.Text, CStr(k)
        End If
        Core_SetToggle "roundsum", True
        AssertEqual 0, Round_RelinkCountForTest(), k & ": ligar o que ja esta ligado nao religa"
        Teardown wb
    Next
End Sub

' Empilhadas com negativo: -1,4 / 2,6 / 2,6 (total 3,8 -> 4; arredondados -1 + 3 + 3 = 5; K = -1; os tres restos sao -0,4: a primeira
' perde a unidade) -> -2 / 3 / 3. Categoria sem valores: total 0, nenhum rotulo e nenhum erro nas Rnd.
Public Sub Test_Round_Stacked_NegativeAndEmptyCategory()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, blk As Range, j As Long
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Fixture(ws, Array(-1.4, 2.6, 2.6, 1, 1, 1)).Select
    ws.Range("B3:D3").ClearContents
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetToggle "roundsum", True
    AssertEqual "-2", cht.SeriesCollection(1).Points(1).DataLabel.Text
    AssertEqual "3", cht.SeriesCollection(2).Points(1).DataLabel.Text
    AssertEqual "3", cht.SeriesCollection(3).Points(1).DataLabel.Text
    AssertEqual "4", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    Set blk = Registry_Block(cht)
    For j = 1 To 3
        AssertFalse IsError(Block_Body(blk).Cells(3, Round_RCol(blk, j)).Value), "categoria vazia: Rnd" & j & " sem erro"
        AssertEqual 0, Block_Body(blk).Cells(3, Round_RCol(blk, j)).Value
        AssertEqual "", cht.SeriesCollection(j).Points(2).DataLabel.Text, "categoria vazia: sem rotulo"
    Next
    Teardown wb
End Sub

' Mudar o formato depois de ligar muda o arredondamento: 1,25 / 1,25 / 1,5 em inteiros 1 / 1 / 2 (total 4); com 1 casa 1,2 / 1,3 / 1,5
' (arredondados 1,3 + 1,3 + 1,5 = 4,1 contra 4,0: K = -1, empate de restos -0,5 entre as duas primeiras: a primeira perde).
Public Sub Test_Round_NumfmtChange_ChangesRounding()
    Dim wb As Workbook, ws As Worksheet, cht As Chart, dec As String
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    Fixture(ws, Array(1.25, 1.25, 1.5, 1, 2, 3)).Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetToggle "roundsum", True
    AssertEqual "1", cht.SeriesCollection(1).Points(1).DataLabel.Text
    AssertEqual "1", cht.SeriesCollection(2).Points(1).DataLabel.Text
    AssertEqual "2", cht.SeriesCollection(3).Points(1).DataLabel.Text
    Core_SetOption "numfmt", "dec1"
    dec = Application.International(xlDecimalSeparator)
    AssertEqual "1" & dec & "2", cht.SeriesCollection(1).Points(1).DataLabel.Text
    AssertEqual "1" & dec & "3", cht.SeriesCollection(2).Points(1).DataLabel.Text
    AssertEqual "1" & dec & "5", cht.SeriesCollection(3).Points(1).DataLabel.Text
    AssertEqual "4" & dec & "0", Chart_FindSeries(cht, "tqq_total").Points(1).DataLabel.Text
    Teardown wb
End Sub
