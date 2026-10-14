Attribute VB_Name = "modChartWaterfallPair"
Option Explicit
' Cascata comparativa (kind waterfallpair, familia waterfallpair, vertical): duas pontes (series 1 e 2 da tabela) lado a lado em cada etapa,
' como o think-cell monta com lacunas de categoria (design da cascata comparativa, secoes 3 e 4). Colunas EMPILHADAS com 3n - 1 categorias
' (fatias): A_i = 3i - 2, B_i = 3i - 1 e o vao_i = 3i (a ultima etapa nao tem vao). Series na ordem tqq_base (invisivel), <serie 1> (parte
' >= 0 de A), tqq_aneg (parte <= 0 de A), <serie 2> (parte >= 0 de B), tqq_bneg: cada fatia so tem valores do seu lado, e base invisivel +
' parte positiva + parte negativa cruzam o zero como a cascata da Fase 1. Categorias de 2 niveis (Outer: nome da etapa na fatia A, vazio na B,
' " " no vao; Inner: " "): o Excel centra o nome sob o par A-B (medido). A linha do eixo de categorias fica invisivel (desenharia os separadores
' do eixo de 2 niveis); a base e a auxiliar XY tqq_zero. Conectores (tqq_conn_a/b), setas de variacao (tqq_pd), sala das setas (tqq_pdroom) e
' nomes das series (tqq_names) sao auxiliares XY no grupo PRINCIPAL (x = numero da fatia, medido) ligadas ao bloco: vivas sem o ThinQQ.
' Rotulos de degrau (tqq_lbl_a/b) e das barras grandes (tqq_lbltot) sao auxiliares de linha. O ThinQQ so recalibra, no refresh, a meia-largura
' das barras (HW), o valor por ponto das setas (PtV), os rotulos finos (LblOn) e "nome dentro da barra ou legenda" (Fit).
' Bloco (corpo, linha 1 = cabecalhos, 6n linhas): etapas (n linhas) 1 Cat | 2..13 lado A | 14..25 lado B (Src Tot D Start End Lo Hi Base PP PN
' Val Big) | 26 ArrowTxt | 27 TotTxtA | 28 TotTxtB; fatias (3n - 1 linhas) 29 Outer .. 43 Top; conectores (n - 1) 44..48; setas (6n - 1:
' 5 pontos por par + linha vazia) 49..51; sala (n) 52..53; nomes (2) 54..56; zero (2) 57..58. Meta (linha 1 do bloco): E PtV, F HW, G Div,
' H Dec, I Suf, J SufD, K Abs, L Lbl, M TotPct, N NameOn, O Fit.
Public Const WFP_COL_CAT As Long = 1
Public Const WFP_SIDE_BASE As Long = 2
Public Const WFP_SIDE_WIDTH As Long = 12
Public Const WFP_COL_ARROWTXT As Long = 26
Public Const WFP_COL_TOTTXT_A As Long = 27
Public Const WFP_COL_TOTTXT_B As Long = 28
Public Const WFP_COL_OUTER As Long = 29
Public Const WFP_COL_INNER As Long = 30
Public Const WFP_COL_SBASE As Long = 31
Public Const WFP_COL_APOS As Long = 32
Public Const WFP_COL_ANEG As Long = 33
Public Const WFP_COL_BPOS As Long = 34
Public Const WFP_COL_BNEG As Long = 35
Public Const WFP_COL_LBLAPOS As Long = 36
Public Const WFP_COL_LBLAVAL As Long = 37
Public Const WFP_COL_LBLBPOS As Long = 38
Public Const WFP_COL_LBLBVAL As Long = 39
Public Const WFP_COL_TOTPOS As Long = 40
Public Const WFP_COL_TOTTXT As Long = 41
Public Const WFP_COL_LBLON As Long = 42
Public Const WFP_COL_TOP As Long = 43
Public Const WFP_COL_CXA As Long = 44
Public Const WFP_COL_CYA As Long = 45
Public Const WFP_COL_CLEN As Long = 46
Public Const WFP_COL_CXB As Long = 47
Public Const WFP_COL_CYB As Long = 48
Public Const WFP_COL_PX As Long = 49
Public Const WFP_COL_PY As Long = 50
Public Const WFP_COL_PLBL As Long = 51
Public Const WFP_COL_RX As Long = 52
Public Const WFP_COL_RY As Long = 53
Public Const WFP_COL_NX As Long = 54
Public Const WFP_COL_NY As Long = 55
Public Const WFP_COL_NTXT As Long = 56
Public Const WFP_COL_ZX As Long = 57
Public Const WFP_COL_ZY As Long = 58
Private Const WFP_COLS As Long = 58
Public Const WFP_META_PTV As Long = 5
Public Const WFP_META_HW As Long = 6
Public Const WFP_META_DIV As Long = 7
Public Const WFP_META_DEC As Long = 8
Public Const WFP_META_SUF As Long = 9
Public Const WFP_META_SUFD As Long = 10
Public Const WFP_META_ABS As Long = 11
Public Const WFP_META_LBL As Long = 12
Public Const WFP_META_TOTPCT As Long = 13
Public Const WFP_META_NAMEON As Long = 14
Public Const WFP_META_FIT As Long = 15
Public Const WFP_ROWS_PER_PAIR As Long = 6       ' linhas do bloco por etapa (as setas tem 5 pontos e uma linha vazia por par)
Public Const WFP_DEFAULT_GAP As Long = 30        ' espaco entre as barras do par (%), o padrao do tipo
Public Const WFP_PT_GAP As Long = 3              ' pt entre o topo da barra e o inicio da seta
Public Const WFP_PT_LBLROOM As Long = 22        ' pt acima de uma barra grande com rotulo: 5,5 de afastamento + 14 do rotulo (medidos) + 2,5
Public Const WFP_PT_RISE As Long = 12            ' pt do inicio mais alto da seta ate a linha horizontal
Public Const WFP_PT_ROOM As Long = 12            ' pt da linha horizontal ate o ponto de sala (meia elipse de 19,9 pt medida + respiro)
Private Const WFP_NAME_MARGIN As Double = 8      ' pt alem da largura estimada do nome (texto vertical)
Private Const WFP_NAME_MIN_W As Double = 14      ' largura (pt) de um texto vertical de 9 pt (medido: 14,0)
Private Const WFP_ERR_FEW As String = "A cascata comparativa precisa de exatamente 2 colunas de valores (uma por série); a tabela tem só 1."
Private Const WFP_ERR_MANY1 As String = "A cascata comparativa compara exatamente 2 séries; a tabela tem "
Private Const WFP_ERR_MANY2 As String = " colunas de valores. Selecione a coluna das etapas e só as duas colunas a comparar."
Private mBusy As Boolean                         ' o refresh grava celulas do bloco: o evento Calculate do grafico reentra em Core_SafeRefresh

Private Sub Wfp_CheckData(data As clsChartData)
    If data.SeriesCount < 2 Then Err.Raise vbObjectError + 6016, "Wfp_Insert", WFP_ERR_FEW
    If data.SeriesCount > 2 Then Err.Raise vbObjectError + 6017, "Wfp_Insert", WFP_ERR_MANY1 & data.SeriesCount & WFP_ERR_MANY2
End Sub

' Indice da parte do lado (0..11) na secao das etapas; -1 para parte desconhecida.
Public Function Wfp_PartIndex(part As String) As Long
    Select Case part
        Case "Src": Wfp_PartIndex = 0
        Case "Tot": Wfp_PartIndex = 1
        Case "D": Wfp_PartIndex = 2
        Case "Start": Wfp_PartIndex = 3
        Case "End": Wfp_PartIndex = 4
        Case "Lo": Wfp_PartIndex = 5
        Case "Hi": Wfp_PartIndex = 6
        Case "Base": Wfp_PartIndex = 7
        Case "PP": Wfp_PartIndex = 8
        Case "PN": Wfp_PartIndex = 9
        Case "Val": Wfp_PartIndex = 10
        Case "Big": Wfp_PartIndex = 11
        Case Else: Wfp_PartIndex = -1
    End Select
End Function

' Coluna (no corpo do bloco) da parte do lado j (1 = serie 1, 2 = serie 2); 0 se o lado ou a parte nao existem.
Public Function Wfp_SideCol(j As Long, part As String) As Long
    Dim k As Long
    k = Wfp_PartIndex(part)
    If k >= 0 And (j = 1 Or j = 2) Then Wfp_SideCol = WFP_SIDE_BASE + WFP_SIDE_WIDTH * (j - 1) + k
End Function

' Numero de fatias (categorias do grafico) para n etapas: A e B por etapa e um vao entre etapas.
Public Function Wfp_SlotCount(n As Long) As Long
    If n > 0 Then Wfp_SlotCount = 3 * n - 1
End Function

' Fatia da barra do lado j (1 = A, 2 = B) na etapa i.
Public Function Wfp_SlotOf(i As Long, j As Long) As Long
    Wfp_SlotOf = 3 * (i - 1) + j
End Function

' Meia-largura da barra em unidades de fatia para o espaco (GapWidth, %): a barra ocupa 1 / (1 + gap/100) da fatia (medido). Pura.
Public Function Wfp_HalfWidth(gapPct As Double) As Double
    Wfp_HalfWidth = 0.5 / (1 + gapPct / 100)
End Function

' Numero de etapas do grafico, pela altura do bloco (meta + cabecalho + 6n). 0 sem bloco.
Public Function Wfp_StepCount(cht As Chart) As Long
    Dim blk As Range
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    Wfp_StepCount = (blk.Rows.Count - 2) \ WFP_ROWS_PER_PAIR
End Function

' j-esima serie NAO auxiliar: 1 = serie 1 (parte >= 0 de A), 2 = serie 2. Nothing se nao existe.
Public Function Wfp_SideSeries(cht As Chart, j As Long) As Series
    Dim s As Series, k As Long
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            k = k + 1
            If k = j Then Set Wfp_SideSeries = s: Exit Function
        End If
    Next
End Function

Private Function SR(body As Range, r As Long, j As Long, part As String) As String
    SR = body.Cells(r, Wfp_SideCol(j, part)).Address(False, False)
End Function

Private Function MetaRef(blk As Range, col As Long) As String
    MetaRef = blk.Cells(1, col).Address(True, True)
End Function

' Constante de matriz com os marcadores de total ({"e","t",...}), a mesma lista da cascata (WF_MARKERS).
Private Function MarkerList() As String
    MarkerList = "{""" & Replace(WF_MARKERS, ",", """,""") & """}"
End Function

' Calculo manual: recalcula o bloco (Range.Calculate segue as dependencias dentro da faixa). Em automatico nao faz nada.
Private Sub Wfp_CalcIfManual(blk As Range)
    If Application.Calculation <> xlCalculationAutomatic Then blk.Calculate
End Sub

' Grava uma celula de meta so quando o valor muda (cada gravacao recalcula e dispara o Calculate do grafico, que reentra no refresh).
Private Sub SetMeta(blk As Range, col As Long, v As Variant)
    Dim cur As Variant
    cur = blk.Cells(1, col).Value
    If VarType(v) = vbString Then
        If Not IsError(cur) Then
            If CStr(cur) = CStr(v) Then Exit Sub
        End If
    ElseIf Not IsError(cur) And Not IsEmpty(cur) And IsNumeric(cur) Then
        If Abs(CDbl(cur) - CDbl(v)) <= 0.000000001 * (1 + Abs(CDbl(v))) Then Exit Sub
    End If
    blk.Cells(1, col).Value = v
    Wfp_CalcIfManual blk
End Sub

Public Function Wfp_BuildBlock(wb As Workbook, id As String, data As clsChartData) As Range
    Dim blk As Range, body As Range, n As Long, ns As Long, i As Long, j As Long, k As Long, r As Long, st As Long
    Dim s As String, t As String, d As String, sta As String, en As String, lo As String, hi As String, bg As String, txt As String
    Dim a As String, b As String, ref1 As String, sA As String, sB As String, yH As String, lblOn As String, srcCat As String
    Dim ptv As String, hw As String, dv As String, dc As String, sf As String, sfd As String, ab As String, lb As String, tp As String, nOn As String, fit As String
    Dim parts As Variant, hdr As Variant, hiMax As Double, loMin As Double, v As Double
    Wfp_CheckData data
    n = data.CategoryCount
    ns = Wfp_SlotCount(n)
    Set blk = CalcSheet_AllocateBlock(wb, id, "waterfallpair", Mid$(ChartData_ExternalAddress(data.Source), 2), WFP_ROWS_PER_PAIR * n + 1, WFP_COLS)
    Set body = Block_Body(blk)
    body.Cells(1, WFP_COL_CAT).Value = "Cat"
    parts = Array("Src", "Tot", "D", "Start", "End", "Lo", "Hi", "Base", "PP", "PN", "Val", "Big")
    For j = 1 To 2
        For k = 0 To 11
            body.Cells(1, Wfp_SideCol(j, CStr(parts(k)))).Value = parts(k) & IIf(j = 1, "A", "B")
        Next
    Next
    hdr = Array("ArrowTxt", "TotTxtA", "TotTxtB", "Outer", "Inner", "SBase", "APos", "ANeg", "BPos", "BNeg", "LblAPos", "LblAVal", "LblBPos", "LblBVal", "TotPos", "TotTxt", "LblOn", "Top", "CxA", "CyA", "CLen", "CxB", "CyB", "PX", "PY", "PLbl", "RX", "RY", "NX", "NY", "NTxt", "ZX", "ZY")
    For k = 0 To UBound(hdr)
        body.Cells(1, WFP_COL_ARROWTXT + k).Value = hdr(k)
    Next
    ' meta: valores iniciais (as formulas abaixo as leem); os textos de sufixo em formato texto ("%" como numero viraria 0%)
    blk.Cells(1, WFP_META_SUF).NumberFormat = "@"
    blk.Cells(1, WFP_META_SUFD).NumberFormat = "@"
    blk.Cells(1, WFP_META_PTV).Value = 1
    blk.Cells(1, WFP_META_HW).Value = Wfp_HalfWidth(WFP_DEFAULT_GAP)
    blk.Cells(1, WFP_META_DIV).Value = 1
    blk.Cells(1, WFP_META_DEC).Value = 0
    blk.Cells(1, WFP_META_ABS).Value = 0
    blk.Cells(1, WFP_META_LBL).Value = 1
    blk.Cells(1, WFP_META_TOTPCT).Value = 1
    blk.Cells(1, WFP_META_NAMEON).Value = 1
    blk.Cells(1, WFP_META_FIT).Value = 1
    ptv = MetaRef(blk, WFP_META_PTV): hw = MetaRef(blk, WFP_META_HW): dv = MetaRef(blk, WFP_META_DIV): dc = MetaRef(blk, WFP_META_DEC)
    sf = MetaRef(blk, WFP_META_SUF): sfd = MetaRef(blk, WFP_META_SUFD): ab = MetaRef(blk, WFP_META_ABS): lb = MetaRef(blk, WFP_META_LBL)
    tp = MetaRef(blk, WFP_META_TOTPCT): nOn = MetaRef(blk, WFP_META_NAMEON): fit = MetaRef(blk, WFP_META_FIT)
    ' etapas: a ponte de cada serie, independente (marcadores por celula)
    For i = 1 To n
        r = i + 1
        srcCat = Mid$(ChartData_ExternalAddress(data.Source.Cells(i + 1, 1)), 2)
        body.Cells(r, WFP_COL_CAT).Formula = "=IF(ISBLANK(" & srcCat & "),"" ""," & srcCat & ")"
        body.Cells(r, WFP_COL_CAT).NumberFormat = data.Source.Cells(i + 1, 1).NumberFormat
        For j = 1 To 2
            s = SR(body, r, j, "Src"): t = SR(body, r, j, "Tot"): d = SR(body, r, j, "D"): sta = SR(body, r, j, "Start")
            en = SR(body, r, j, "End"): lo = SR(body, r, j, "Lo"): hi = SR(body, r, j, "Hi"): bg = SR(body, r, j, "Big")
            body.Cells(r, Wfp_SideCol(j, "Src")).Formula = ChartData_ExternalAddress(data.ValueCell(i, j))
            body.Cells(r, Wfp_SideCol(j, "Tot")).Formula = "=IF(ISERROR(" & s & "),FALSE,IF(ISNUMBER(" & s & "),FALSE,OR(LOWER(TRIM(" & s & "&""""))=" & MarkerList() & ")))"
            body.Cells(r, Wfp_SideCol(j, "D")).Formula = "=IF(" & t & ",0,IF(ISNUMBER(" & s & ")," & s & ",0))"
            If i = 1 Then
                body.Cells(r, Wfp_SideCol(j, "Start")).Formula = "=0"
            Else
                body.Cells(r, Wfp_SideCol(j, "Start")).Formula = "=" & SR(body, r - 1, j, "End")
            End If
            body.Cells(r, Wfp_SideCol(j, "End")).Formula = "=IF(" & t & "," & sta & "," & sta & "+" & d & ")"
            body.Cells(r, Wfp_SideCol(j, "Lo")).Formula = "=IF(" & t & ",MIN(0," & en & "),MIN(" & sta & "," & en & "))"
            body.Cells(r, Wfp_SideCol(j, "Hi")).Formula = "=IF(" & t & ",MAX(0," & en & "),MAX(" & sta & "," & en & "))"
            body.Cells(r, Wfp_SideCol(j, "Base")).Formula = "=MAX(0," & lo & ")+MIN(0," & hi & ")"
            body.Cells(r, Wfp_SideCol(j, "PP")).Formula = "=MAX(0," & hi & ")-MAX(0," & lo & ")"
            body.Cells(r, Wfp_SideCol(j, "PN")).Formula = "=MIN(0," & lo & ")-MIN(0," & hi & ")"
            If i = 1 Then
                body.Cells(r, Wfp_SideCol(j, "Big")).Formula = "=TRUE"
            Else
                body.Cells(r, Wfp_SideCol(j, "Big")).Formula = "=" & t
            End If
            body.Cells(r, Wfp_SideCol(j, "Val")).Formula = "=IF(" & bg & "," & en & "," & d & ")"
            ' texto das barras grandes: o valor (FIXED: separadores do Excel que calculou o bloco por ultimo) e, depois da 1a etapa, o percentual do 1o total da serie
            ref1 = body.Cells(2, Wfp_SideCol(j, "End")).Address(True, True)
            If i = 1 Then
                txt = "=IF(" & bg & ",FIXED(ROUND(" & en & "/" & dv & "," & dc & ")," & dc & ")&" & sf & ","""")"
            Else
                txt = "=IF(" & bg & ",FIXED(ROUND(" & en & "/" & dv & "," & dc & ")," & dc & ")&" & sf & "&IF(AND(" & tp & "=1," & ref1 & "<>0),"" (""&FIXED(ROUND(100*" & en & "/" & ref1 & ",0),0)&""%)"",""""),"""")"
            End If
            If j = 1 Then body.Cells(r, WFP_COL_TOTTXT_A).Formula = txt Else body.Cells(r, WFP_COL_TOTTXT_B).Formula = txt
        Next
        ' texto da seta: (B - A) / A em % inteiro com sinal; "n/d" com A = 0 ou sinais opostos; com Abs = 1, B - A no formato do grafico
        a = SR(body, r, 1, "Val"): b = SR(body, r, 2, "Val")
        body.Cells(r, WFP_COL_ARROWTXT).Formula = "=IF(" & ab & "=1,IF(ROUND((" & b & "-" & a & ")/" & dv & "," & dc & ")>0,""+"","""")&FIXED(ROUND((" & b & "-" & a & ")/" & dv & "," & dc & ")," & dc & ")&" & sfd & ",IF(OR(" & a & "=0," & a & "*" & b & "<0),""n/d"",IF(ROUND(100*(" & b & "-" & a & ")/" & a & ",0)>0,""+"","""")&FIXED(ROUND(100*(" & b & "-" & a & ")/" & a & ",0),0)&""%""))"
    Next
    ' fatias: o que o grafico plota
    For i = 1 To n
        st = i + 1
        For j = 1 To 2
            k = Wfp_SlotOf(i, j): r = k + 1
            bg = SR(body, st, j, "Big"): lo = SR(body, st, j, "Lo"): hi = SR(body, st, j, "Hi"): en = SR(body, st, j, "End"): d = SR(body, st, j, "D")
            lblOn = body.Cells(r, WFP_COL_LBLON).Address(False, False)
            If j = 1 Then
                body.Cells(r, WFP_COL_OUTER).Formula = "=" & body.Cells(st, WFP_COL_CAT).Address(False, False)
                body.Cells(r, WFP_COL_OUTER).NumberFormat = data.Source.Cells(i + 1, 1).NumberFormat
            End If
            body.Cells(r, WFP_COL_INNER).Value = " "
            body.Cells(r, WFP_COL_SBASE).Formula = "=" & SR(body, st, j, "Base")
            If j = 1 Then
                body.Cells(r, WFP_COL_APOS).Formula = "=" & SR(body, st, 1, "PP")
                body.Cells(r, WFP_COL_ANEG).Formula = "=" & SR(body, st, 1, "PN")
                body.Cells(r, WFP_COL_BPOS).Formula = "=0"
                body.Cells(r, WFP_COL_BNEG).Formula = "=0"
                body.Cells(r, WFP_COL_LBLAPOS).Formula = "=IF(" & bg & ",NA(),(" & lo & "+" & hi & ")/2)"
                body.Cells(r, WFP_COL_LBLAVAL).Formula = "=IF(OR(" & bg & "," & lblOn & "=0),""""," & d & ")"
                body.Cells(r, WFP_COL_LBLBPOS).Formula = "=NA()"
                body.Cells(r, WFP_COL_LBLBVAL).Formula = "="""""
                body.Cells(r, WFP_COL_TOTTXT).Formula = "=" & body.Cells(st, WFP_COL_TOTTXT_A).Address(False, False)
            Else
                body.Cells(r, WFP_COL_APOS).Formula = "=0"
                body.Cells(r, WFP_COL_ANEG).Formula = "=0"
                body.Cells(r, WFP_COL_BPOS).Formula = "=" & SR(body, st, 2, "PP")
                body.Cells(r, WFP_COL_BNEG).Formula = "=" & SR(body, st, 2, "PN")
                body.Cells(r, WFP_COL_LBLAPOS).Formula = "=NA()"
                body.Cells(r, WFP_COL_LBLAVAL).Formula = "="""""
                body.Cells(r, WFP_COL_LBLBPOS).Formula = "=IF(" & bg & ",NA(),(" & lo & "+" & hi & ")/2)"
                body.Cells(r, WFP_COL_LBLBVAL).Formula = "=IF(OR(" & bg & "," & lblOn & "=0),""""," & d & ")"
                body.Cells(r, WFP_COL_TOTTXT).Formula = "=" & body.Cells(st, WFP_COL_TOTTXT_B).Address(False, False)
            End If
            body.Cells(r, WFP_COL_TOTPOS).Formula = "=IF(" & bg & ",IF(" & en & ">=0," & hi & "," & lo & "),NA())"
            body.Cells(r, WFP_COL_LBLON).Value = 1
            body.Cells(r, WFP_COL_TOP).Formula = "=" & hi
        Next
        If i < n Then
            r = 3 * i + 1                                  ' vao da etapa i (fatia 3i)
            body.Cells(r, WFP_COL_OUTER).Value = " "
            body.Cells(r, WFP_COL_INNER).Value = " "
            body.Cells(r, WFP_COL_SBASE).Resize(1, 5).Formula = "=0"
            body.Cells(r, WFP_COL_LBLAPOS).Formula = "=NA()"
            body.Cells(r, WFP_COL_LBLAVAL).Formula = "="""""
            body.Cells(r, WFP_COL_LBLBPOS).Formula = "=NA()"
            body.Cells(r, WFP_COL_LBLBVAL).Formula = "="""""
            body.Cells(r, WFP_COL_TOTPOS).Formula = "=NA()"
            body.Cells(r, WFP_COL_TOTTXT).Formula = "="""""
            body.Cells(r, WFP_COL_LBLON).Value = 1
            body.Cells(r, WFP_COL_TOP).Formula = "=0"
        End If
    Next
    ' conectores: do lado direito da barra i ao lado esquerdo da barra i + 1 da MESMA serie, na altura do fim da barra i
    For i = 1 To n - 1
        r = i + 1
        body.Cells(r, WFP_COL_CXA).Formula = "=" & (3 * i - 2) & "+" & hw
        body.Cells(r, WFP_COL_CYA).Formula = "=" & SR(body, r, 1, "End")
        body.Cells(r, WFP_COL_CLEN).Formula = "=3-2*" & hw
        body.Cells(r, WFP_COL_CXB).Formula = "=" & (3 * i - 1) & "+" & hw
        body.Cells(r, WFP_COL_CYB).Formula = "=" & SR(body, r, 2, "End")
    Next
    ' setas: 5 pontos por par (sobe de A, horizontal, desce em B) e uma linha vazia; o afastamento em pontos vezes PtV (valor por ponto)
    For i = 1 To n
        st = i + 1
        sA = "(" & SR(body, st, 1, "Hi") & "+" & ptv & "*IF(AND(" & SR(body, st, 1, "Big") & "," & lb & "=1," & SR(body, st, 1, "End") & ">=0)," & WFP_PT_LBLROOM & "," & WFP_PT_GAP & "))"
        sB = "(" & SR(body, st, 2, "Hi") & "+" & ptv & "*IF(AND(" & SR(body, st, 2, "Big") & "," & lb & "=1," & SR(body, st, 2, "End") & ">=0)," & WFP_PT_LBLROOM & "," & WFP_PT_GAP & "))"
        yH = "(MAX(" & sA & "," & sB & ")+" & ptv & "*" & WFP_PT_RISE & ")"
        r = WFP_ROWS_PER_PAIR * (i - 1) + 2
        body.Cells(r, WFP_COL_PX).Formula = "=" & (3 * i - 2): body.Cells(r, WFP_COL_PY).Formula = "=" & sA
        body.Cells(r + 1, WFP_COL_PX).Formula = "=" & (3 * i - 2): body.Cells(r + 1, WFP_COL_PY).Formula = "=" & yH
        body.Cells(r + 2, WFP_COL_PX).Formula = "=" & (6 * i - 3) & "/2": body.Cells(r + 2, WFP_COL_PY).Formula = "=" & yH
        body.Cells(r + 2, WFP_COL_PLBL).Formula = "=" & body.Cells(st, WFP_COL_ARROWTXT).Address(False, False)
        body.Cells(r + 3, WFP_COL_PX).Formula = "=" & (3 * i - 1): body.Cells(r + 3, WFP_COL_PY).Formula = "=" & yH
        body.Cells(r + 4, WFP_COL_PX).Formula = "=" & (3 * i - 1): body.Cells(r + 4, WFP_COL_PY).Formula = "=" & sB
        body.Cells(st, WFP_COL_RX).Formula = "=" & (6 * i - 3) & "/2"
        body.Cells(st, WFP_COL_RY).Formula = "=" & yH & "+" & ptv & "*" & WFP_PT_ROOM
    Next
    ' nomes das series: meio da 1a barra de cada lado; texto so com NameOn = 1 e Fit = 1
    For j = 1 To 2
        r = j + 1
        body.Cells(r, WFP_COL_NX).Formula = "=" & j
        body.Cells(r, WFP_COL_NY).Formula = "=(" & SR(body, 2, j, "Lo") & "+" & SR(body, 2, j, "Hi") & ")/2"
        body.Cells(r, WFP_COL_NTXT).Formula = "=IF(AND(" & nOn & "=1," & fit & "=1)," & Mid$(ChartData_ExternalAddress(data.SeriesNameCell(j)), 2) & "&"""","""")"
    Next
    ' linha do zero: de 0,5 a 3n - 0,5 (bordas da area interna)
    body.Cells(2, WFP_COL_ZX).Formula = "=1/2": body.Cells(2, WFP_COL_ZY).Formula = "=0"
    body.Cells(3, WFP_COL_ZX).Formula = "=" & (2 * ns + 1) & "/2": body.Cells(3, WFP_COL_ZY).Formula = "=0"
    ' PtV inicial: a faixa dos dados em 180 pt (o refresh calibra pelo eixo real)
    Wfp_CalcIfManual blk
    hiMax = Application.WorksheetFunction.Max(body.Cells(2, WFP_COL_TOP).Resize(ns, 1))
    For i = 1 To n
        For j = 1 To 2
            v = Util_ToDouble(body.Cells(i + 1, Wfp_SideCol(j, "Lo")).Value)
            If v < loMin Then loMin = v
        Next
    Next
    v = (hiMax - loMin) / 180
    If v <= 0 Then v = 1
    blk.Cells(1, WFP_META_PTV).Value = v
    Wfp_CalcIfManual blk
    Set Wfp_BuildBlock = blk
End Function

Private Function AddBar(cht As Chart, name As String, valueCells As Range, cats As String) As Series
    Dim s As Series
    Set s = cht.SeriesCollection.NewSeries
    s.Name = name
    s.Values = Block_Ref(valueCells)
    s.XValues = cats
    Set AddBar = s
End Function

' Auxiliar XY no grupo PRINCIPAL (x = numero da fatia, medido), sem marcador; withLine: linha ligando os pontos (setas, zero). Em falha desfaz.
Private Function AddXY(cht As Chart, name As String, xCells As Range, yCells As Range, withLine As Boolean) As Series
    Dim s As Series, en As Long, es As String, ed As String
    On Error GoTo Fail
    Set s = cht.SeriesCollection.NewSeries
    s.Name = name
    If withLine Then s.ChartType = xlXYScatterLinesNoMarkers Else s.ChartType = xlXYScatter     ' antes de XValues: XY aceita X numerico
    s.XValues = Block_Ref(xCells)
    s.Values = Block_Ref(yCells)
    s.MarkerStyle = xlMarkerStyleNone
    If Not withLine Then s.Format.Line.Visible = msoFalse
    Set AddXY = s
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error GoTo -1                                 ' encerra o tratamento do erro: sem isso o Resume Next abaixo nao protegeria a limpeza
    On Error Resume Next
    If Not s Is Nothing Then s.Delete
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

Private Sub Wfp_StyleCategoryAxis(cht As Chart)
    With cht.Axes(xlCategory)
        .TickLabelPosition = xlTickLabelPositionLow
        .TickLabels.MultiLevel = True
        .TickLabels.Orientation = xlHorizontal
        .Format.Line.Visible = msoFalse                ' a linha do eixo desenharia os separadores do eixo de 2 niveis (medido)
    End With
End Sub

Private Sub Wfp_AddZeroLine(cht As Chart)
    Dim body As Range, s As Series
    If Not Chart_FindSeries(cht, "tqq_zero") Is Nothing Then Exit Sub
    Set body = Block_Body(Registry_Block(cht))
    Set s = AddXY(cht, "tqq_zero", body.Cells(2, WFP_COL_ZX).Resize(2, 1), body.Cells(2, WFP_COL_ZY).Resize(2, 1), True)
    s.Format.Line.Visible = msoTrue
    s.Format.Line.ForeColor.RGB = RGB(166, 166, 166)
    s.Format.Line.Weight = 0.75
End Sub

' Nomes das series: auxiliar XY tqq_names (meio da 1a barra de cada lado) com o rotulo ligado a NTxt, vertical, 9 pt, fonte por contraste.
Private Sub Wfp_SetNames(cht As Chart)
    Dim body As Range, s As Series, k As Long
    Set body = Block_Body(Registry_Block(cht))
    Set s = Chart_FindSeries(cht, "tqq_names")
    If s Is Nothing Then Set s = AddXY(cht, "tqq_names", body.Cells(2, WFP_COL_NX).Resize(2, 1), body.Cells(2, WFP_COL_NY).Resize(2, 1), False)
    For k = 1 To 2
        With s.Points(k)
            .HasDataLabel = True
            .DataLabel.Formula = Block_Ref(body.Cells(k + 1, WFP_COL_NTXT))
            .DataLabel.Position = xlLabelPositionCenter
            .DataLabel.Orientation = xlUpward
            .DataLabel.Font.Size = 9
            .DataLabel.Font.Bold = False
            .DataLabel.Font.Color = Style_ContrastFont(Style_DataSeriesColor(cht, k))
        End With
    Next
End Sub

' Conectores: duas auxiliares XY sem marcador (tqq_conn_a, tqq_conn_b), um ponto por conector (n - 1) no lado direito da barra i, com barra de
' erro X custom (CLen do bloco) ate o lado esquerdo da barra i + 1 da MESMA serie, tracejada e sem ponta (medido; vivos sem o ThinQQ). Passam
' sobre a barra da outra serie (como no think-cell). Uma etapa so: sem conectores (a opcao fica gravada).
' Faixa da barra de erro em R1C1 (a forma que o Excel aceita para valores personalizados; em A1 pode recusar com 1004 ou virar constante).
' Transacional: a opcao so e gravada depois das series prontas; em falha as duas somem, a opcao fica "0" (o estado real) e o erro e repropagado.
Public Sub Wfp_SetConnectors(cht As Chart, turnOn As Boolean)
    Dim body As Range, n As Long, j As Long, s As Series, nm As String, en As Long, es As String, ed As String
    DropConnectors cht
    If turnOn Then
        n = Wfp_StepCount(cht)
        If n >= 2 Then
            On Error GoTo Fail
            Set body = Block_Body(Registry_Block(cht))
            For j = 1 To 2
                nm = "tqq_conn_" & Mid$("ab", j, 1)
                If j = 1 Then
                    Set s = AddXY(cht, nm, body.Cells(2, WFP_COL_CXA).Resize(n - 1, 1), body.Cells(2, WFP_COL_CYA).Resize(n - 1, 1), False)
                Else
                    Set s = AddXY(cht, nm, body.Cells(2, WFP_COL_CXB).Resize(n - 1, 1), body.Cells(2, WFP_COL_CYB).Resize(n - 1, 1), False)
                End If
                s.ErrorBar Direction:=xlX, Include:=xlErrorBarIncludePlusValues, Type:=xlErrorBarTypeCustom, Amount:="='" & CALC_SHEET & "'!" & body.Cells(2, WFP_COL_CLEN).Resize(n - 1, 1).Address(True, True, xlR1C1)
                With s.ErrorBars
                    .EndStyle = xlNoCap
                    .Format.Line.Visible = msoTrue
                    .Format.Line.ForeColor.RGB = RGB(127, 127, 127)
                    .Format.Line.DashStyle = msoLineDash
                    .Format.Line.Weight = 0.75
                End With
            Next
            On Error GoTo 0
            If cht.HasLegend Then Chart_HideHelperLegendEntries cht
        End If
    End If
    Registry_SetOption cht, "connectors", IIf(turnOn, "1", "0")
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    DropConnectors cht
    Registry_SetOption cht, "connectors", "0"
    On Error GoTo 0
    Err.Raise en, es, ed
End Sub

Private Sub DropConnectors(cht As Chart)
    Dim j As Long, s As Series
    For j = 1 To 2
        Set s = Chart_FindSeries(cht, "tqq_conn_" & Mid$("ab", j, 1))
        If Not s Is Nothing Then s.Delete
    Next
End Sub

Public Function Wfp_Insert(ws As Worksheet, data As clsChartData) As ChartObject
    Dim id As String, co As ChartObject, cht As Chart, anchor As Range, blk As Range, body As Range, n As Long, ns As Long, cats As String
    Dim en As Long, es As String, ed As String
    Wfp_CheckData data                               ' recusa (6016/6017) antes de alocar: nada a desfazer
    On Error GoTo Fail
    id = Registry_NewId()
    Set blk = Wfp_BuildBlock(ws.Parent, id, data)
    Util_TestHookAfterBlock
    Set body = Block_Body(blk)
    n = data.CategoryCount
    ns = Wfp_SlotCount(n)
    cats = Block_Ref(body.Cells(2, WFP_COL_OUTER).Resize(ns, 2))     ' 2 colunas: eixo de categorias de 2 niveis
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, Chart_FreeTop(ws, anchor.Left, anchor.Top, 460, 280), 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    AddBar cht, "tqq_base", body.Cells(2, WFP_COL_SBASE).Resize(ns, 1), cats
    AddBar cht, ChartData_ExternalAddress(data.SeriesNameCell(1)), body.Cells(2, WFP_COL_APOS).Resize(ns, 1), cats
    AddBar cht, "tqq_aneg", body.Cells(2, WFP_COL_ANEG).Resize(ns, 1), cats
    AddBar cht, ChartData_ExternalAddress(data.SeriesNameCell(2)), body.Cells(2, WFP_COL_BPOS).Resize(ns, 1), cats
    AddBar cht, "tqq_bneg", body.Cells(2, WFP_COL_BNEG).Resize(ns, 1), cats
    cht.ChartType = xlColumnStacked
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlNotPlotted               ' as linhas vazias do bloco separam as setas; as colunas sao formulas (nunca vazias)
    Style_ApplyBase cht
    Util_TestHookAfterChart
    cht.ChartGroups(1).GapWidth = WFP_DEFAULT_GAP
    Wfp_StyleCategoryAxis cht
    cht.HasLegend = False
    Registry_SetOption cht, "numfmt", "int"
    Registry_SetOption cht, "palette", "padrao"
    Registry_SetOption cht, "gap", CStr(WFP_DEFAULT_GAP)
    Registry_SetOption cht, "n", CStr(n)
    Registry_SetOption cht, "totals", "1"
    Registry_SetOption cht, "legend", "inbar"
    Wfp_ApplyColors cht, "padrao"
    Wfp_AddZeroLine cht
    Wfp_SetNames cht
    Wfp_SetConnectors cht, True
    Wfp_SetLegend cht, "inbar"
    Wfp_RefreshAfterDataChange cht
    Set Wfp_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error GoTo -1                                 ' encerra o tratamento do erro: sem isso o Resume Next abaixo nao protegeria a limpeza
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' Cores: serie 1 = pal(0) e serie 2 = pal(2) (o contraste escuro/claro da imagem do think-cell), as partes negativas com a cor da serie, a base
' invisivel; a fonte dos rotulos e dos nomes acompanha (contraste).
Public Sub Wfp_ApplyColors(cht As Chart, paletteName As String)
    Dim pal As Variant, s As Series, k As Long, c As Variant
    pal = Style_Palette(paletteName)
    For Each s In cht.SeriesCollection
        c = Empty
        Select Case s.Name
            Case "tqq_base"
                s.Format.Fill.Visible = msoFalse
                s.Format.Line.Visible = msoFalse
            Case "tqq_aneg"
                c = pal(0)
            Case "tqq_bneg"
                c = pal(2)
            Case Else
                If Not Style_IsHelperSeries(s) Then
                    k = k + 1
                    If k = 1 Then c = pal(0) Else c = pal(2)
                End If
        End Select
        If Not IsEmpty(c) Then
            s.Format.Fill.Visible = msoTrue
            Style_SetColor s.Format.Fill.ForeColor, c
            s.Format.Line.Visible = msoFalse
        End If
    Next
    Wfp_RefreshFontColors cht
End Sub

Private Function PointHasLabel(s As Series, i As Long) As Boolean
    On Error Resume Next
    PointHasLabel = s.Points(i).HasDataLabel
End Function

' Fonte por contraste com a cor atual de cada serie: rotulos de degrau (ponto a ponto: uma escrita no nivel da serie recriaria rotulos apagados)
' e o nome de cada serie.
Private Sub Wfp_RefreshFontColors(cht As Chart)
    Dim s As Series, k As Long, i As Long, fc As Long
    For k = 1 To 2
        fc = Style_ContrastFont(Style_DataSeriesColor(cht, k))
        If k = 1 Then Set s = Chart_FindSeries(cht, "tqq_lbl_a") Else Set s = Chart_FindSeries(cht, "tqq_lbl_b")
        If Not s Is Nothing Then
            For i = 1 To s.Points.Count
                If PointHasLabel(s, i) Then s.Points(i).DataLabel.Font.Color = fc
            Next
        End If
        Set s = Chart_FindSeries(cht, "tqq_names")
        If Not s Is Nothing Then
            If s.Points.Count >= k Then
                If PointHasLabel(s, k) Then s.Points(k).DataLabel.Font.Color = fc
            End If
        End If
    Next
End Sub

Private Sub ShowLegend(cht As Chart, turnOn As Boolean)
    If turnOn Then
        If Not cht.HasLegend Then cht.HasLegend = True
        cht.Legend.Position = xlLegendPositionRight
        cht.Legend.Font.Size = 9
        Chart_HideHelperLegendEntries cht
    ElseIf cht.HasLegend Then
        cht.HasLegend = False
    End If
End Sub

' Legenda do tipo: right (legenda a direita, sem nomes nas barras), none (nada), inbar (nomes verticais na 1a barra; se algum nao cabe, os dois
' somem e a legenda aparece: Wfp_SyncNames). Quem chama grava a opcao "legend" antes (Wfp_SyncNames a le).
Public Sub Wfp_SetLegend(cht As Chart, mode As String)
    Dim blk As Range
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    SetMeta blk, WFP_META_NAMEON, IIf(mode = "inbar", 1, 0)
    Select Case mode
        Case "right": ShowLegend cht, True
        Case "none": ShowLegend cht, False
        Case Else: Wfp_SyncNames cht
    End Select
End Sub

' "Nome dentro das barras": o nome cabe se a 1a barra da serie tem altura >= largura estimada do nome + 8 pt e largura >= 14 pt (texto
' vertical de 9 pt, medido). Se algum nao cabe, Fit = 0 (os dois nomes ficam vazios) e a legenda aparece; voltando a caber, o contrario (a
' legenda so muda quando o estado dela nao corresponde: legenda visivel = algum nome nao cabe). So com a opcao inbar; grava so o que muda.
Public Sub Wfp_SyncNames(cht As Chart)
    Dim blk As Range, body As Range, ax As Axis, ppu As Double, barW As Double, j As Long, h As Double, need As Double, fits As Boolean, s As Series, ns As Long
    If Registry_GetOption(cht, "legend", "inbar") <> "inbar" Then Exit Sub
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    Set body = Block_Body(blk)
    ns = Wfp_SlotCount(Wfp_StepCount(cht))
    Set ax = cht.Axes(xlValue, xlPrimary)
    If ns < 1 Or ax.MaximumScale <= ax.MinimumScale Then Exit Sub
    ppu = cht.PlotArea.InsideHeight / (ax.MaximumScale - ax.MinimumScale)
    barW = cht.PlotArea.InsideWidth / ns / (1 + cht.ChartGroups(1).GapWidth / 100)
    fits = (barW >= WFP_NAME_MIN_W)
    For j = 1 To 2
        Set s = Wfp_SideSeries(cht, j)
        h = (Util_ToDouble(body.Cells(2, Wfp_SideCol(j, "Hi")).Value) - Util_ToDouble(body.Cells(2, Wfp_SideCol(j, "Lo")).Value)) * ppu
        need = WFP_NAME_MARGIN
        If Not s Is Nothing Then need = need + Labels_EstimateWidth(s.Name)
        If h < need Then fits = False
    Next
    SetMeta blk, WFP_META_NAMEON, 1
    SetMeta blk, WFP_META_FIT, IIf(fits, 1, 0)
    If cht.HasLegend = fits Then ShowLegend cht, Not fits
End Sub

' Meia-largura das barras (meta HW) a partir do espaco REAL do grupo de colunas (mudado pela faixa ou a mao no Excel). Grava so se mudou.
Public Sub Wfp_SyncGap(cht As Chart)
    Dim blk As Range
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    SetMeta blk, WFP_META_HW, Wfp_HalfWidth(CDbl(cht.ChartGroups(1).GapWidth))
End Sub

' O que depende dos DADOS, do tamanho ou do espaco e muda depois de inserir (Core_SafeRefresh): meia-largura, nomes/legenda. Guarda de
' reentrada (as gravacoes disparam o Calculate do grafico). Cada passo e isolado (a falha de um e registrada e nao pula os seguintes). Nunca
' propaga erro.
Public Sub Wfp_RefreshAfterDataChange(cht As Chart)
    On Error GoTo Done
    If mBusy Then Exit Sub
    If Kinds_Family(Registry_ChartType(cht)) <> "waterfallpair" Then Exit Sub
    mBusy = True
    On Error Resume Next
    Wfp_SyncGap cht
    If Err.Number <> 0 Then Errors_Log "Wfp_RefreshAfterDataChange/gap" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Wfp_SyncNames cht
    If Err.Number <> 0 Then Errors_Log "Wfp_RefreshAfterDataChange/nomes" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
Done:
    If Err.Number <> 0 Then Errors_Log "Wfp_RefreshAfterDataChange" & vbTab & Err.Number & vbTab & Err.Description
    mBusy = False
End Sub

' Linhas da tabela de origem entre a 1a e a ultima etapa (pelas referencias Src do lado 1 no bloco): linhas inseridas/apagadas no meio da
' tabela mudam esse numero (as series apontam para o bloco, como no Mekko). -1 se nao for possivel ler.
Public Function Wfp_SourceRowsSpanned(cht As Chart) As Long
    Dim body As Range, n As Long, r1 As Long, r2 As Long
    On Error GoTo Fail
    Wfp_SourceRowsSpanned = -1
    n = Wfp_StepCount(cht)
    If n < 1 Then Exit Function
    Set body = Block_Body(Registry_Block(cht))
    r1 = RefRow(CStr(body.Cells(2, Wfp_SideCol(1, "Src")).Formula))
    r2 = RefRow(CStr(body.Cells(n + 1, Wfp_SideCol(1, "Src")).Formula))
    If r1 < 1 Or r2 < r1 Then Exit Function
    Wfp_SourceRowsSpanned = r2 - r1 + 1
    Exit Function
Fail:
    Wfp_SourceRowsSpanned = -1
End Function

' Linha da referencia ='Planilha'!$B$7 (o numero depois do ultimo "$"); 0 sem referencia (#REF!).
Private Function RefRow(f As String) As Long
    Dim p As Long
    If InStr(f, "#REF!") > 0 Then Exit Function
    p = InStrRev(f, "$")
    If p = 0 Then Exit Function
    RefRow = CLng(Val(Mid$(f, p + 1)))
End Function
