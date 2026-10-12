Attribute VB_Name = "modRounding"
Option Explicit
' Rotulos com soma consistente (opcao roundsum): arredondamento por MAIOR RESTO em formulas nativas do bloco (sem UDF: uma funcao de
' suplemento vira vinculo externo ao caminho do .xlam quando a pasta e aberta sem o ThinQQ, e o caminho muda a cada versao do core).
' Para cada categoria (linha), os S valores da linha sao arredondados a D casas e a diferenca para o alvo arredondado (nos 100%, a soma
' das participacoes: 1, ou 0 numa linha vazia; a coluna Total nos demais) e distribuida, uma unidade por vez, aos valores de maior resto
' (ou retirada dos de menor resto), com empate desfeito pela ordem das colunas (a PRIMEIRA coluna empatada recebe a unidade, ou a perde).
' So ROUND, SUMPRODUCT, POWER e IF (sem COUNTIF, sem entrada matricial): verificado no Excel 2019.
' As colunas RndD | V_1..V_S | Rnd_1..Rnd_S sao acrescentadas A DIREITA do bloco na primeira vez (Block_Extend); V_j = valor saneado
' (empilhadas: IF(ISNUMBER(src),src,0), com src lido da faixa da linha na formula do Total; 100%: =Part_j; area: =Val_j). Os rotulos de
' segmento passam a apontar para Rnd_j; o rotulo do total continua ligado ao Total, que o formato mostra como ROUND(Total, D) = soma dos Rnd_j.
' D (RndD) = casas do formato dos SEGMENTOS: nos 100% a participacao (pct0, ou pct1 quando escolhido; Labels_SegmentFormat), nos demais numfmt.
' O cabecalho RndD e gravado POR ULTIMO: ele marca as colunas como prontas (uma falha no meio desfaz a extensao e nada fica pela metade).
Private Const RND_HDR_D As String = "RndD"
Private Const RND_ERR_RANGE As String = "Selecione uma única coluna ou linha só com números (pelo menos 2)."
Private Const AGG_PREFIX As String = "=AGGREGATE(9,6,"
Public Const ROUND_MAX_CELLS As Long = 2000            ' Arredondar intervalo: limite de celulas (custo O(n^2) por recalculo)

Public Function Round_Decimals(numfmtKey As String) As Long
    Select Case numfmtKey
        Case "dec1": Round_Decimals = 1
        Case "dec2": Round_Decimals = 2
        Case "pct0": Round_Decimals = 2
        Case "pct1": Round_Decimals = 3
        Case "mil": Round_Decimals = -3
        Case "mi": Round_Decimals = -5
        Case Else: Round_Decimals = 0
    End Select
End Function

' Casas dos rotulos de segmento do grafico: nos 100% eles mostram a participacao (pct0, ou pct1 quando o usuario escolhe pct1), como em
' Labels_SegmentFormat; nos demais, o numfmt.
Private Function SegmentDecimals(cht As Chart) As Long
    Dim key As String
    key = Registry_GetOption(cht, "numfmt", "int")
    If Kinds_Is100(Registry_ChartType(cht)) Then key = IIf(key = "pct1", "pct1", "pct0")
    SegmentDecimals = Round_Decimals(key)
End Function

' Formula (texto em ingles, com "=") do maior resto para UMA celula: v = endereco do valor, rowRange = os S valores da linha,
' partialRange = do primeiro valor ate v (empate pela ordem), target = alvo (endereco ou expressao), decimals = endereco ou literal.
Public Function Round_Formula(v As String, rowRange As String, partialRange As String, target As String, decimals As String) As String
    Dim u As String, k As String, remExpr As String, rd As String, ra As String
    u = "POWER(10,-" & decimals & ")"
    k = "ROUND((ROUND(" & target & "," & decimals & ")-SUMPRODUCT(ROUND(" & rowRange & "," & decimals & ")))/" & u & ",0)"
    remExpr = "(" & v & "-ROUND(" & v & "," & decimals & "))/" & u
    rd = "(SUMPRODUCT(--((" & rowRange & "-ROUND(" & rowRange & "," & decimals & "))/" & u & ">" & remExpr & "))+SUMPRODUCT(--((" & partialRange & "-ROUND(" & partialRange & "," & decimals & "))/" & u & "=" & remExpr & ")))"
    ra = "(SUMPRODUCT(--((" & rowRange & "-ROUND(" & rowRange & "," & decimals & "))/" & u & "<" & remExpr & "))+SUMPRODUCT(--((" & partialRange & "-ROUND(" & partialRange & "," & decimals & "))/" & u & "=" & remExpr & ")))"
    Round_Formula = "=ROUND(" & v & "," & decimals & ")+" & u & "*(IF(" & k & ">0,--(" & rd & "<=" & k & "),0)-IF(" & k & "<0,--(" & ra & "<=-" & k & "),0))"
End Function

Public Function Round_SeriesCount(cht As Chart) As Long
    Dim s As Series
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then Round_SeriesCount = Round_SeriesCount + 1
    Next
End Function

Public Function Round_DCol(blk As Range) As Long
    Round_DCol = Block_FindCol(blk, RND_HDR_D)
End Function

Public Function Round_VCol(blk As Range, j As Long) As Long
    Round_VCol = Block_FindCol(blk, "V" & j)
End Function

Public Function Round_RCol(blk As Range, j As Long) As Long
    Round_RCol = Block_FindCol(blk, "Rnd" & j)
End Function

' Numero de colunas Rnd_j do bloco (Rnd1, Rnd2, ... seguidas); 0 sem as colunas.
Private Function RoundColumnCount(blk As Range) As Long
    Do While Round_RCol(blk, RoundColumnCount + 1) > 0
        RoundColumnCount = RoundColumnCount + 1
    Loop
End Function

' Faixa (texto com a planilha) dos valores de origem da linha r do corpo de um bloco empilhado, lida da formula do Total
' (=AGGREGATE(9,6,<faixa>)): o Excel a mantem viva quando a tabela ganha ou perde linhas acima ou a planilha muda de nome, ao contrario do
' endereco de origem gravado na meta. "" se a formula nao tem essa forma.
Private Function StackedRowRef(body As Range, r As Long) As String
    Dim f As String
    f = CStr(body.Cells(r, 2).Formula)
    If Left$(f, Len(AGG_PREFIX)) <> AGG_PREFIX Or Right$(f, 1) <> ")" Or Len(f) <= Len(AGG_PREFIX) + 1 Then Exit Function
    StackedRowRef = Mid$(f, Len(AGG_PREFIX) + 1, Len(f) - Len(AGG_PREFIX) - 1)
End Function

' Gancho de teste (Util_FailThinSyncForTest = "round-<ponto>:<chave>", como o scatter-/bfly-): falha simulada num ponto do codigo para testar
' o caminho de erro. Dispara UMA vez (limpa o gancho antes de levantar o erro: a limpeza do caminho de erro passa pelo mesmo ponto).
' Inerte fora do build Test.
Private Sub RoundTestHook(site As String, key As String)
    If BUILD_CONFIG = "Test" Then
        If Len(Util_FailThinSyncForTest) > 0 Then
            If Util_FailThinSyncForTest = "round-" & site & ":" & key Then
                Util_FailThinSyncForTest = ""
                Err.Raise vbObjectError + 9995, "ThinQQ", "Falha simulada (teste) " & site & " " & key
            End If
        End If
    End If
End Sub

' Cria as colunas (uma vez) e grava as formulas. Devolve o bloco (ja estendido). Transacional: uma falha depois de estender limpa as
' colunas novas e devolve o nome tqq_<id> a faixa anterior antes de o erro seguir para o chamador.
Public Function Round_EnsureColumns(cht As Chart) As Range
    Dim blk As Range, body As Range, wb As Workbook, id As String, kind As String, fam As String, is100 As Boolean, ns As Long, n As Long
    Dim c0 As Long, i As Long, j As Long, r As Long, src As Range, refs() As String, f() As Variant, d As Long, w As Long
    Dim vAddr As String, rowRng As String, partRng As String, target As String, dAddr As String, oldRef As String, extended As Boolean
    Dim en As Long, es As String, ed As String
    Set blk = Registry_Block(cht)
    id = Registry_IdFromChart(cht)
    If blk Is Nothing Then Err.Raise vbObjectError + 6015, "Round_EnsureColumns", "Bloco inexistente: " & id
    If Round_DCol(blk) > 0 Then Set Round_EnsureColumns = blk: Exit Function
    Set wb = Registry_Workbook(cht)
    kind = Registry_ChartType(cht): fam = Kinds_Family(kind): is100 = Kinds_Is100(kind)
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    ' S = colunas que o alvo soma: as do bloco na area (Val_j) e nos 100% (Part_j); nas empilhadas, as series de dados do grafico
    If fam = "area" Then
        ns = Area_SeriesCount(blk)
    ElseIf is100 Then
        Do While Block_FindCol(blk, "Part" & (ns + 1)) > 0
            ns = ns + 1
        Loop
    Else
        ns = Round_SeriesCount(cht)
    End If
    If ns < 1 Or n < 1 Then Err.Raise vbObjectError + 7002, "Round_EnsureColumns", "O gráfico não tem séries de dados."
    If fam = "stacked" And Not is100 Then
        ReDim refs(1 To n)
        For i = 1 To n
            refs(i) = StackedRowRef(body, i + 1)
            If Len(refs(i)) = 0 Then
                If src Is Nothing Then Set src = Block_SourceRange(cht)
                If src Is Nothing Then Err.Raise vbObjectError + 7002, "Round_EnsureColumns", "Não foi possível ler a tabela de origem do gráfico. Use 'Alterar dados' para reconstruí-lo."
                refs(i) = Mid$(ChartData_ExternalAddress(src.Cells(i + 1, 2).Resize(1, ns)), 2)
            End If
        Next
    End If
    d = SegmentDecimals(cht)
    c0 = blk.Columns.Count
    w = 2 * ns + 1
    oldRef = wb.Names(ID_PREFIX & id).RefersTo
    On Error GoTo Fail
    Set blk = Block_Extend(wb, id, w)
    extended = True
    Set body = Block_Body(blk)
    ' tudo numa gravacao so (cada gravacao de celula recalcula e dispara os eventos dos graficos); RndD no cabecalho so no fim
    ReDim f(1 To n + 1, 1 To w)
    f(1, 1) = ""
    For j = 1 To ns
        f(1, 1 + j) = "V" & j
        f(1, 1 + ns + j) = "Rnd" & j
    Next
    For i = 1 To n
        r = i + 1
        f(r, 1) = d
        For j = 1 To ns
            If fam = "area" Then
                f(r, 1 + j) = "=" & body.Cells(r, Area_ValCol(blk, j)).Address(False, False)
            ElseIf is100 Then
                f(r, 1 + j) = "=" & body.Cells(r, 3 + j).Address(False, False)
            Else
                f(r, 1 + j) = "=IF(ISNUMBER(INDEX(" & refs(i) & ",1," & j & ")),INDEX(" & refs(i) & ",1," & j & "),0)"
            End If
        Next
        rowRng = body.Cells(r, c0 + 2).Resize(1, ns).Address(False, False)
        dAddr = body.Cells(r, c0 + 1).Address(False, False)
        If is100 Then target = "SUM(" & rowRng & ")" Else target = body.Cells(r, 2).Address(False, False)     ' 100%: 1, ou 0 numa linha vazia (sem "1%" fantasma)
        For j = 1 To ns
            vAddr = body.Cells(r, c0 + 1 + j).Address(False, False)
            partRng = body.Cells(r, c0 + 2).Resize(1, j).Address(False, False)
            f(r, 1 + ns + j) = Round_Formula(vAddr, rowRng, partRng, target, dAddr)
        Next
    Next
    body.Cells(1, c0 + 1).Resize(n + 1, w).Formula = f
    RoundTestHook "ensure", "write"
    body.Cells(1, c0 + 1).Value = RND_HDR_D
    Round_ApplyCellFormat cht
    Set Round_EnsureColumns = blk
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error GoTo -1                                 ' encerra o tratamento do erro: sem isso o Resume Next abaixo nao protegeria a limpeza
    On Error Resume Next
    If extended Then
        blk.Columns(c0 + 1).Resize(blk.Rows.Count, w).Clear
        wb.Names(ID_PREFIX & id).RefersTo = oldRef
    End If
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' RndD = casas do formato atual dos segmentos; celulas Rnd_j com o formato dos segmentos (o rotulo ligado mostra o texto da celula). Sem as
' colunas nao faz nada. So grava RndD quando muda (cada gravacao recalcula o bloco).
Public Sub Round_ApplyCellFormat(cht As Chart)
    Dim blk As Range, body As Range, n As Long, dcol As Long, j As Long, rc As Long, d As Long, i As Long, fmt As String
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    dcol = Round_DCol(blk)
    If dcol = 0 Then Exit Sub
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    If n < 1 Then Exit Sub
    d = SegmentDecimals(cht)
    For i = 2 To n + 1
        If CStr(body.Cells(i, dcol).Formula) <> CStr(d) Then body.Cells(2, dcol).Resize(n, 1).Value = d: Exit For
    Next
    fmt = Style_LocalizeCode(Labels_SegmentFormat(cht))
    For j = 1 To RoundColumnCount(blk)
        rc = Round_RCol(blk, j)
        body.Cells(2, rc).Resize(n, 1).NumberFormatLocal = fmt
    Next
End Sub

' Rotulos de segmento refeitos pela familia (Stacked_SetLabels/Area_SetLabels leem a opcao roundsum).
Private Sub RelinkSegments(cht As Chart, fam As String)
    Dim s As Series
    If fam = "area" Then
        Area_SetLabels cht, True
    Else
        ' rotulos nativos nao se desfazem de uma ligacao por ponto ao receber propriedades no nivel da serie: recomeca do zero
        For Each s In cht.SeriesCollection
            If Not Style_IsHelperSeries(s) Then s.HasDataLabels = False
        Next
        RoundTestHook "relink", "cleared"
        Stacked_SetLabels cht, True
    End If
End Sub

' Liga/desliga a opcao e religa os rotulos de segmento pela familia. Recusa (7002) fora da matriz. Desligar nao apaga as colunas (ficam
' calculadas; religar nao estende de novo); desligar o que nao esta ligado so grava "0" (nao refaz os rotulos). Com os rotulos desligados so
' grava a opcao (ligar os rotulos depois liga as Rnd). Transacional: uma falha ao religar volta a opcao anterior e refaz os rotulos com ela
' antes de o erro seguir para o chamador.
Public Sub Round_SetLabels(cht As Chart, turnOn As Boolean)
    Dim kind As String, fam As String, prev As String, recorded As Boolean, en As Long, es As String, ed As String
    kind = Registry_ChartType(cht): fam = Kinds_Family(kind)
    If Not Kinds_Supports(kind, "roundsum") Then Err.Raise vbObjectError + 7002, "Round_SetLabels", Kinds_NotApplicableText("roundsum")
    prev = Registry_GetOption(cht, "roundsum", "0")
    If Not turnOn And prev <> "1" Then Registry_SetOption cht, "roundsum", "0": Exit Sub
    If turnOn Then Round_EnsureColumns cht                   ' desfaz a propria extensao se falhar; nada gravado
    On Error GoTo Fail
    Registry_SetOption cht, "roundsum", IIf(turnOn, "1", "0")
    recorded = True
    If Registry_GetOption(cht, "labels", "1") <> "1" Then Exit Sub
    RelinkSegments cht, fam
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    If recorded Then
        Registry_SetOption cht, "roundsum", prev
        If Registry_GetOption(cht, "labels", "1") = "1" Then RelinkSegments cht, fam
        If Err.Number <> 0 Then Errors_Log "Round_SetLabels/desfazer" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    End If
    On Error GoTo 0
    Err.Raise en, es, ed
End Sub

' "" se a faixa serve ao Arredondar intervalo (uma area, uma coluna OU uma linha, de 2 a ROUND_MAX_CELLS celulas, todas numericas); senao a
' mensagem. O limite vem do custo: cada formula percorre a faixa inteira varias vezes (O(n) por celula, O(n^2) por recalculo).
Public Function Round_ValidateRange(rng As Range) As String
    Dim c As Range
    Round_ValidateRange = RND_ERR_RANGE
    If rng Is Nothing Then Exit Function
    If rng.Areas.Count <> 1 Then Exit Function
    If rng.Rows.Count > 1 And rng.Columns.Count > 1 Then Exit Function
    If rng.Cells.CountLarge < 2 Then Exit Function
    If rng.Cells.CountLarge > ROUND_MAX_CELLS Then Round_ValidateRange = Round_TooBigText(rng.Cells.CountLarge): Exit Function
    For Each c In rng.Cells
        If IsError(c.Value) Or IsEmpty(c.Value) Or VarType(c.Value) = vbString Or VarType(c.Value) = vbBoolean Then Exit Function
        If Not IsNumeric(c.Value) Then Exit Function
    Next
    Round_ValidateRange = ""
End Function

Public Function Round_TooBigText(n As Variant) As String
    Round_TooBigText = "A faixa tem " & Format$(n, "#,##0") & " células; o Arredondar intervalo aceita no máximo " & Format$(ROUND_MAX_CELLS, "#,##0") & " (cada fórmula percorre a faixa inteira e o recálculo cresceria com o quadrado do tamanho)."
End Function

' Celulas de saida para a faixa src a partir de dest (a forma da origem); Nothing se nao cabem na planilha.
Private Function OutputRange(src As Range, dest As Range) As Range
    Dim n As Long
    n = src.Cells.Count
    If src.Columns.Count = 1 Then
        If dest.Row + n - 1 > dest.Worksheet.Rows.Count Then Exit Function
        Set OutputRange = dest.Cells(1, 1).Resize(n, 1)
    Else
        If dest.Column + n - 1 > dest.Worksheet.Columns.Count Then Exit Function
        Set OutputRange = dest.Cells(1, 1).Resize(1, n)
    End If
End Function

' "" se a celula de destino serve para as formulas da faixa src (mesma pasta, a saida inteira cabe na planilha, nao sobrepoe a origem, sem
' celula mesclada, que esconderia parte do resultado, e sem celula bloqueada numa planilha protegida); senao a mensagem.
Public Function Round_ValidateDest(src As Range, dest As Range) As String
    Dim out As Range, flag As Variant
    If dest Is Nothing Then Exit Function
    If Not dest.Worksheet.Parent Is src.Worksheet.Parent Then Round_ValidateDest = "Escolha a célula de destino na mesma pasta de trabalho da faixa de origem.": Exit Function
    Set out = OutputRange(src, dest)
    If out Is Nothing Then Round_ValidateDest = "O resultado não cabe na planilha a partir dessa célula.": Exit Function
    If out.Worksheet Is src.Worksheet Then
        If Not Application.Intersect(out, src) Is Nothing Then Round_ValidateDest = "O resultado não pode ficar sobre a faixa de origem.": Exit Function
    End If
    flag = out.MergeCells                                       ' Null = algumas mescladas
    If IsNull(flag) Then flag = True
    If flag Then Round_ValidateDest = "O destino tem células mescladas; escolha uma área sem mesclagem.": Exit Function
    If out.Worksheet.ProtectContents Then
        flag = out.Locked                                       ' Null = algumas bloqueadas
        If IsNull(flag) Then flag = True
        If flag Then Round_ValidateDest = "As células de destino estão protegidas.": Exit Function
    End If
End Function

' Grava em dest (celula inicial; a saida tem a forma da origem) uma formula por celula que referencia SO a faixa de origem: portatil.
' Valida antes (tamanho e destino: erro 7002 com a mensagem, nada gravado). Uma gravacao so (matriz em Range.Formula), com a tela congelada
' e o calculo manual; tela e modo de calculo voltam aos valores anteriores em toda saida. Transacional: uma falha depois de comecar a
' gravar devolve as formulas/valores e os formatos anteriores das celulas de saida antes de o erro seguir.
Public Sub Round_WriteFormulas(src As Range, dest As Range, decimals As Long)
    Dim n As Long, i As Long, vertical As Boolean, rowRng As String, partRng As String, out As Range, fmt As String, msg As String
    Dim f() As Variant, prevF As Variant, prevFmt() As String, prevCalc As Long, prevSU As Boolean, started As Boolean
    Dim en As Long, es As String, ed As String
    If src.Cells.CountLarge > ROUND_MAX_CELLS Then Err.Raise vbObjectError + 7002, "Round_WriteFormulas", Round_TooBigText(src.Cells.CountLarge)
    msg = Round_ValidateDest(src, dest)
    If Len(msg) > 0 Then Err.Raise vbObjectError + 7002, "Round_WriteFormulas", msg
    n = src.Cells.Count
    vertical = (src.Columns.Count = 1)
    Set out = OutputRange(src, dest)
    rowRng = Mid$(ChartData_ExternalAddress(src), 2)
    If vertical Then ReDim f(1 To n, 1 To 1) Else ReDim f(1 To 1, 1 To n)
    For i = 1 To n
        If vertical Then
            partRng = Mid$(ChartData_ExternalAddress(src.Cells(1, 1).Resize(i, 1)), 2)
            f(i, 1) = Round_Formula(Mid$(ChartData_ExternalAddress(src.Cells(i)), 2), rowRng, partRng, "SUM(" & rowRng & ")", "(" & decimals & ")")
        Else
            partRng = Mid$(ChartData_ExternalAddress(src.Cells(1, 1).Resize(1, i)), 2)
            f(1, i) = Round_Formula(Mid$(ChartData_ExternalAddress(src.Cells(i)), 2), rowRng, partRng, "SUM(" & rowRng & ")", "(" & decimals & ")")
        End If
    Next
    If decimals > 0 Then fmt = "#,##0." & String$(decimals, "0") Else fmt = "#,##0"
    prevF = out.Formula
    ReDim prevFmt(1 To n)
    For i = 1 To n: prevFmt(i) = out.Cells(i).NumberFormat: Next
    prevSU = Application.ScreenUpdating
    prevCalc = Application.Calculation
    On Error GoTo Fail
    started = True
    Application.ScreenUpdating = False
    If prevCalc <> xlCalculationManual Then Application.Calculation = xlCalculationManual
    out.Formula = f
    RoundTestHook "write", "formulas"
    out.NumberFormatLocal = Style_LocalizeCode(fmt)
    out.Calculate                                               ' em calculo manual (o do usuario, inclusive) o resultado ja sai calculado
    If Application.Calculation <> prevCalc Then Application.Calculation = prevCalc
    If Application.ScreenUpdating <> prevSU Then Application.ScreenUpdating = prevSU
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error GoTo -1                                 ' encerra o tratamento do erro: sem isso o Resume Next abaixo nao protegeria a limpeza
    On Error Resume Next
    If started Then
        out.Formula = prevF
        For i = 1 To n: out.Cells(i).NumberFormat = prevFmt(i): Next
    End If
    If Application.Calculation <> prevCalc Then Application.Calculation = prevCalc
    If Application.ScreenUpdating <> prevSU Then Application.ScreenUpdating = prevSU
    If Err.Number <> 0 Then Errors_Log "Round_WriteFormulas/desfazer" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    On Error GoTo 0
    Err.Raise en, es, ed
End Sub
