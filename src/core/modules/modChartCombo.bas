Attribute VB_Name = "modChartCombo"
Option Explicit
' Combinacao: colunas empilhadas (as series 1..S-1 da tabela) + uma linha com marcadores (a ultima coluna da tabela), com
' eixo secundario opcional para a linha (kind combo). O bloco e o da familia empilhada (Stacked_BuildBlock com kind combo:
' Total, TotalPos e VLine so das colunas) e totais, rotulos, conectores, cores e o numero de categorias vem dos Stacked_*.
Private Const CB_LBL_H As Double = 14           ' altura de um rotulo de 9 pt numa linha (medido: 13,99 pt; total e linha)
Private Const CB_LBL_OFF As Double = 5.5        ' distancia do ponto ao rotulo acima dele (medido: 5,49 pt no total e na linha)
Private Const CB_MARKER As Double = 2.5         ' raio do marcador da linha (MarkerSize 5)
Private Const CB_GAP As Double = 4             ' respiro entre o marcador da linha e o topo do rotulo do total
Private Const CB_BAND_FRAC As Double = 0.12    ' altura da faixa da linha: 12% da area interna...
Private Const CB_BAND_MIN As Double = 24       ' ... e no minimo 24 pt
Private Const CB_BAND_CAP As Double = 0.25     ' ... e no maximo 25% (a faixa que o minimo 0 pede; acima disso a linha usa a faixa toda)
Private Const CB_MAX_TOPFRAC As Double = 0.5   ' a faixa de cima (linha, rotulos e respiros) ocupa no maximo metade da area interna
Private Const CB_KEEP_UP As Double = 0.002     ' histerese do maximo primario: subir menos que 0,2% nao regrava...
Private Const CB_KEEP_DOWN As Double = 0.01    ' ... nem descer menos que 1% (sobra folga)
Private Const CB_KEEP_SEC As Double = 0.005    ' histerese do secundario: 0,5% da faixa do eixo

' Exige pelo menos 2 colunas de valores (uma de colunas e a da linha). Erro de validacao dos dados (6008): Core_Insert mostra a
' mensagem como aviso, sem registrar falha; nada e criado. O eixo secundario nasce ligado quando as magnitudes diferem muito
' (Combo_DefaultSecondary). Em falha depois da insercao desfaz o grafico e o bloco.
Public Function Combo_Insert(ws As Worksheet, data As clsChartData) As ChartObject
    Dim co As ChartObject, id As String, en As Long, es As String, ed As String
    If data.SeriesCount < 2 Then
        Err.Raise vbObjectError + 6008, "Combo_Insert", "A combinação precisa de pelo menos 2 colunas de valores: as primeiras viram colunas empilhadas e a última vira linha."
    End If
    On Error GoTo Fail
    Set co = Stacked_Insert(ws, data, "combo")      ' desfaz o que criou em caso de falha
    Util_TestHookAfterChart                          ' (teste) falha com o grafico ja criado: o Fail: abaixo o desfaz
    Combo_SetSecondary co.Chart, Combo_DefaultSecondary(data)
    Set Combo_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then
        id = Mid$(co.Name, Len(ID_PREFIX) + 1)
        co.Delete
        If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    End If
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' Eixo secundario por padrao quando as magnitudes diferem por mais de 5 vezes (a linha de percentuais sobre colunas de
' milhares ficaria achatada no eixo comum): maior |valor| da ultima coluna (linha) x maior |soma das demais| por linha
' (colunas). Valores lidos por Util_ToDouble (texto/vazio = 0). Falso se alguma das magnitudes for zero.
Public Function Combo_DefaultSecondary(data As clsChartData) As Boolean
    Dim i As Long, j As Long, ns As Long, lineMax As Double, colMax As Double, v As Double, rowSum As Double
    ns = data.SeriesCount
    If ns < 2 Then Exit Function
    For i = 1 To data.CategoryCount
        v = Abs(Util_ToDouble(data.ValueCell(i, ns).Value))
        If v > lineMax Then lineMax = v
        rowSum = 0
        For j = 1 To ns - 1
            rowSum = rowSum + Util_ToDouble(data.ValueCell(i, j).Value)
        Next
        If Abs(rowSum) > colMax Then colMax = Abs(rowSum)
    Next
    If lineMax > 0 And colMax > 0 Then Combo_DefaultSecondary = (lineMax > 5 * colMax Or lineMax * 5 < colMax)
End Function

' Liga ou desliga o eixo secundario da linha. O eixo de valores secundario fica oculto como o primario (a escala ainda existe
' para a geometria). So em combinacao (nos demais tipos nao faz nada nem grava a opcao). A mudanca de grupo pode devolver as
' entradas das auxiliares a legenda: reescondidas no fim.
Public Sub Combo_SetSecondary(cht As Chart, turnOn As Boolean)
    Dim ln As Series
    If Not Kinds_Supports(Registry_ChartType(cht), "secondary") Then Exit Sub
    Set ln = Combo_LineSeries(cht)
    If ln Is Nothing Then Exit Sub
    If turnOn Then
        ln.AxisGroup = xlSecondary
        cht.HasAxis(xlValue, xlSecondary) = True
        Style_HideValueAxis cht.Axes(xlValue, xlSecondary)
    Else
        ln.AxisGroup = xlPrimary
    End If
    Chart_HideHelperLegendEntries cht
    Registry_SetOption cht, "secondary", IIf(turnOn, "1", "0")
    Combo_SyncSecondaryScale cht
End Sub

' Escala do eixo secundario (oculto) e folga das colunas. Com TOTAIS e todos os valores da linha >= 0 (FitAboveTotals): a linha ocupa uma
' faixa no alto da area de plotagem, ACIMA do maior rotulo de total (o padrao do Excel punha a linha sobre os totais: medido na demonstracao,
' o rotulo do total mais alto ja comeca 5 pt ACIMA da area interna, entao nenhuma escala secundaria sozinha a tiraria de cima deles). Para
' isso o maximo do eixo PRIMARIO (oculto) sobe o bastante para as colunas deixarem a faixa livre (opcao "primax"; so com o maximo automatico
' e sem Mesma escala) e o secundario poe o maior valor da linha no topo da faixa e o menor no fundo (minimo 0 quando cabe, como na Fase 2).
' Sem totais (ou sem caber, ou com valor negativo na linha): a regra da Fase 2, minimo 0 com todos os valores >= 0 e automatico com algum
' negativo. As escalas fixadas aqui ficam gravadas ("secmin", "secmax", "primax"): um minimo ou maximo FIXO diferente do gravado foi
' definido a mao pelo usuario e nao e sobrescrito (a folga das colunas sai). Uma passada, sem laco: a area de plotagem nao depende das
' escalas (eixos de valores ocultos), entao o refresh seguinte chega aos mesmos numeros. So escreve o que mudou. Sem eixo secundario
' devolve a folga das colunas. Nunca propaga erro (chamada por eventos).
Public Sub Combo_SyncSecondaryScale(cht As Chart)
    Dim ln As Series, v As Variant, i As Long, nonNeg As Boolean, hasNum As Boolean, ax As Axis, lmin As Double, lmax As Double, x As Double
    Dim fitted As Boolean
    On Error GoTo Done
    If Registry_ChartType(cht) <> "combo" Then Exit Sub
    If Registry_GetOption(cht, "secondary", "0") <> "1" Then ReleasePrimary cht: Exit Sub
    If Not cht.HasAxis(xlValue, xlSecondary) Then ReleasePrimary cht: Exit Sub
    Set ln = Combo_LineSeries(cht)
    If ln Is Nothing Then Exit Sub
    nonNeg = True
    v = ln.Values
    If IsArray(v) Then
        For i = LBound(v) To UBound(v)
            If Not IsError(v(i)) And Not IsEmpty(v(i)) Then
                If VarType(v(i)) <> vbString And VarType(v(i)) <> vbBoolean Then
                    If IsNumeric(v(i)) Then
                        x = CDbl(v(i))
                        If x < 0 Then nonNeg = False
                        If Not hasNum Then
                            lmin = x: lmax = x: hasNum = True
                        Else
                            If x < lmin Then lmin = x
                            If x > lmax Then lmax = x
                        End If
                    End If
                End If
            End If
        Next
    End If
    Set ax = cht.Axes(xlValue, xlSecondary)
    If Not ax.MinimumScaleIsAuto And Not OptIs(cht, "secmin", ax.MinimumScale) Then ReleasePrimary cht: GoTo Done   ' minimo do usuario: fica
    If Not ax.MaximumScaleIsAuto And Not OptIs(cht, "secmax", ax.MaximumScale) Then ReleasePrimary cht: GoTo Done   ' maximo do usuario: fica
    If nonNeg And hasNum Then fitted = FitAboveTotals(cht, ax, lmin, lmax)
    If fitted Then GoTo Done
    ReleasePrimary cht
    If Not ax.MaximumScaleIsAuto Then ax.MaximumScaleIsAuto = True                   ' (so pode ser o gravado: o do usuario saiu acima)
    If Len(Registry_GetOption(cht, "secmax", "")) > 0 Then Registry_SetOption cht, "secmax", ""
    If nonNeg Then
        If ax.MinimumScaleIsAuto Or ax.MinimumScale <> 0 Then ax.MinimumScale = 0
        If Registry_GetOption(cht, "secmin", "") <> "0" Then Registry_SetOption cht, "secmin", "0"
    Else
        If Not ax.MinimumScaleIsAuto Then ax.MinimumScaleIsAuto = True
        If Len(Registry_GetOption(cht, "secmin", "")) > 0 Then Registry_SetOption cht, "secmin", ""
    End If
Done:
    If Err.Number <> 0 Then Errors_Log "Combo_SyncSecondaryScale" & vbTab & Err.Number & vbTab & Err.Description
End Sub

Private Function SameNum(a As Double, b As Double) As Boolean
    SameNum = (Abs(a - b) <= 0.000000001 * (1 + Abs(a) + Abs(b)))
End Function

' A opcao key guarda (Str$, ponto decimal) exatamente o valor v? (reconhece a escala fixada pelo proprio ThinQQ)
Private Function OptIs(cht As Chart, key As String, v As Double) As Boolean
    Dim s As String
    s = Registry_GetOption(cht, key, "")
    If Len(s) = 0 Then Exit Function
    OptIs = SameNum(Val(s), v)
End Function

Private Function NumText(v As Double) As String
    NumText = Trim$(Str$(v))
End Function

' Maximo do eixo primario de volta ao automatico se for o da folga das colunas (primax) e limpa a opcao; um maximo do usuario (ou da Mesma
' escala) fica como esta.
Private Sub ReleasePrimary(cht As Chart)
    Dim pax As Axis
    If Len(Registry_GetOption(cht, "primax", "")) = 0 Then Exit Sub
    Set pax = cht.Axes(xlValue, xlPrimary)
    If Not pax.MaximumScaleIsAuto And OptIs(cht, "primax", pax.MaximumScale) And Len(Registry_GetOption(cht, "vmax", "")) = 0 Then pax.MaximumScaleIsAuto = True
    Registry_SetOption cht, "primax", ""
End Sub

' Minimo e maximo de um eixo, na ordem que o Excel aceita (o minimo novo nunca passa do maximo atual), so se mudaram.
Private Sub SetScale(ax As Axis, mn As Double, mx As Double)
    If Not ax.MinimumScaleIsAuto And Not ax.MaximumScaleIsAuto Then
        If SameNum(ax.MinimumScale, mn) And SameNum(ax.MaximumScale, mx) Then Exit Sub
    End If
    If mn < ax.MaximumScale Then
        ax.MinimumScale = mn: ax.MaximumScale = mx
    Else
        ax.MaximumScale = mx: ax.MinimumScale = mn
    End If
End Sub

' A linha numa faixa acima do maior rotulo de total (ver Combo_SyncSecondaryScale). Geometria (pt, coordenadas do grafico): t e h = topo e
' altura da area interna; o topo da faixa yTop deixa o rotulo da linha (acima do ponto) dentro do grafico; a faixa tem CB_BAND_FRAC da area
' interna (no minimo CB_BAND_MIN); abaixo dela o marcador, um respiro e o rotulo do total (altura e distancia ao ponto medidas: 14 e 5,5 pt)
' ate o topo da coluna mais alta, yc. Se a faixa de cima passaria de CB_MAX_TOPFRAC da area interna (grafico baixo demais), nada muda.
' Maximo primario m: colMax (TotalPos, a posicao plotada do total) em yc = t + h (m - colMax) / (m - pmin) => m = (colMax - f pmin) / (1 - f),
' f = (yc - t) / h; o automatico fica se ja deixa a folga. Depois, com a escala primaria REAL, a faixa vai de yTop ao fundo livre e o
' secundario poe lmax no topo e lmin no fundo: minimo 0 quando lmin / lmax >= (1 - b) / (1 - a) (a, b = fracoes de topo e fundo da faixa),
' senao a faixa inteira (r = (lmax - lmin) / (b - a)). False se nao cabe (ex.: Mesma escala fixou as colunas no alto) ou nao ha totais.
Private Function FitAboveTotals(cht As Chart, ax As Axis, lmin As Double, lmax As Double) As Boolean
    Dim tot As Series, tv As Variant, i As Long, colMax As Double, t As Double, h As Double, yTop As Double, band As Double, yc As Double
    Dim f As Double, m As Double, pax As Axis, pmn As Double, pmx As Double, yBot As Double, a As Double, b As Double, smin As Double
    Dim smax As Double, r As Double, below As Double, free As Boolean, b0 As Double
    If Registry_GetOption(cht, "totals", "0") <> "1" Then Exit Function
    Set tot = Chart_FindSeries(cht, "tqq_total")
    If tot Is Nothing Then Exit Function
    If lmax <= 0 Then Exit Function                                    ' linha toda em zero: nada a separar
    tv = tot.Values
    If Not IsArray(tv) Then Exit Function
    For i = LBound(tv) To UBound(tv)
        If Not IsError(tv(i)) And Not IsEmpty(tv(i)) Then
            If IsNumeric(tv(i)) Then
                If CDbl(tv(i)) > colMax Then colMax = CDbl(tv(i))
            End If
        End If
    Next
    If colMax <= 0 Then Exit Function
    t = cht.PlotArea.InsideTop: h = cht.PlotArea.InsideHeight
    If h <= 0 Then Exit Function
    below = CB_MARKER + CB_GAP + CB_LBL_H + CB_LBL_OFF
    yTop = CB_LBL_H + CB_LBL_OFF + 2
    If yTop < t Then yTop = t
    a = (yTop - t) / h
    ' faixa: a que o minimo 0 pede (lmin no fundo: b0 = 1 - (1 - a) lmin / lmax), entre o piso (CB_BAND_FRAC, CB_BAND_MIN) e o teto
    ' CB_BAND_CAP; acima do teto, o piso (e a linha ocupa a faixa inteira, com o minimo abaixo de 0)
    band = h * CB_BAND_FRAC
    If band < CB_BAND_MIN Then band = CB_BAND_MIN
    b0 = (1 - (1 - a) * lmin / lmax - a) * h
    If b0 > band And b0 <= h * CB_BAND_CAP Then band = b0
    yc = yTop + band + below
    f = (yc - t) / h
    If f > CB_MAX_TOPFRAC Then Exit Function
    Set pax = cht.Axes(xlValue, xlPrimary)
    free = (Len(Registry_GetOption(cht, "vmax", "")) = 0 And Len(Registry_GetOption(cht, "vmin", "")) = 0)
    If free And Not pax.MaximumScaleIsAuto Then free = OptIs(cht, "primax", pax.MaximumScale)
    If free Then
        pmn = pax.MinimumScale
        m = (colMax - f * pmn) / (1 - f)
        If pax.MaximumScaleIsAuto Then
            If pax.MaximumScale < m Then
                pax.MaximumScale = m
                Registry_SetOption cht, "primax", NumText(m)
            ElseIf Len(Registry_GetOption(cht, "primax", "")) > 0 Then
                Registry_SetOption cht, "primax", ""                     ' o automatico ja deixa a folga
            End If
        ElseIf m > pax.MaximumScale * (1 + CB_KEEP_UP) Or m < pax.MaximumScale * (1 - CB_KEEP_DOWN) Then
            ' histerese: o layout automatico do Excel muda a area interna um pouco quando as colunas mudam (medido: 59.400 -> 59.430 no
            ' refresh seguinte); um ajuste menor que isso cabe nos respiros e nao e regravado (o refresh seguinte nao mexe de novo)
            pax.MaximumScale = m
            Registry_SetOption cht, "primax", NumText(m)
        End If
    ElseIf Len(Registry_GetOption(cht, "primax", "")) > 0 Then
        Registry_SetOption cht, "primax", ""                             ' Mesma escala ou o usuario tomou o eixo: a folga deixa de ser nossa
    End If
    pmn = pax.MinimumScale: pmx = pax.MaximumScale
    If pmx <= pmn Then Exit Function
    yBot = t + h * (pmx - colMax) / (pmx - pmn) - below
    If yBot - yTop < CB_BAND_MIN / 2 Then Exit Function
    b = (yBot - t) / h
    If lmin / lmax >= (1 - b) / (1 - a) - CB_KEEP_DOWN Then             ' minimo 0 (tolerancia: ate 1% da area abaixo do fundo, nos respiros)
        smin = 0: smax = lmax / (1 - a)
    Else
        r = (lmax - lmin) / (b - a)
        smax = lmax + a * r: smin = smax - r
    End If
    ' histerese tambem no secundario: uma diferenca de ate 0,5% da faixa do eixo nao e regravada
    If Not ax.MinimumScaleIsAuto And Not ax.MaximumScaleIsAuto Then
        If Abs(ax.MinimumScale - smin) <= CB_KEEP_SEC * (smax - smin) And Abs(ax.MaximumScale - smax) <= CB_KEEP_SEC * (smax - smin) Then
            smin = ax.MinimumScale: smax = ax.MaximumScale
        End If
    End If
    SetScale ax, smin, smax
    If Not OptIs(cht, "secmin", smin) Then Registry_SetOption cht, "secmin", NumText(smin)
    If Not OptIs(cht, "secmax", smax) Then Registry_SetOption cht, "secmax", NumText(smax)
    FitAboveTotals = True
End Function

' O que numa combinacao depende dos DADOS e muda depois de inserir (Core_SafeRefresh): a escala do eixo secundario e o formato
' dos rotulos da linha (que segue o formato da celula). Nunca propaga erro.
Public Sub Combo_RefreshAfterDataChange(cht As Chart)
    On Error Resume Next
    If Registry_ChartType(cht) <> "combo" Then Exit Sub
    Combo_SyncSecondaryScale cht
    Labels_RefreshComboLineFormat cht
End Sub

' A serie de linha da combinacao: a ultima serie nao auxiliar (Nothing se nao ha nenhuma).
Public Function Combo_LineSeries(cht As Chart) As Series
    Dim s As Series
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then Set Combo_LineSeries = s
    Next
End Function
