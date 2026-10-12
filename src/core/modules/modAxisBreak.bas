Attribute VB_Name = "modAxisBreak"
Option Explicit
' Corte de eixo (opcao vbreak; design da Fase 3, secao 4.6): o MAXIMO FIXO do eixo de valores corta as colunas/barras na borda da area de
' plotagem (medido: o Excel desenha ate a borda); o rotulo do total passa a usar a coluna LblPos = MIN(TotalPos, BreakAt) do bloco (vivo),
' os rotulos dos segmentos acima do corte sao ocultados e o do segmento atravessado vai para a base (Break_SyncLabels), e cada coluna
' cortada recebe uma marca de corte (3 linhas: branca por baixo e duas cinzas; formas tqq_ann_break_<i>, posicionadas por Ann_Refresh).
' Sem o ThinQQ: eixo fixo e LblPos sao vivos; rotulos de segmento e marcas ficam como na ultima sincronizacao. Mesma escala e corte sao
' excludentes (Break_Set limpa vmin/vmax; Deco_SameScale ignora graficos com corte).
' Rotulos ocultos pelo corte ficam na opcao "brklbl" (pares serie.ponto, como o "thinlbl" da cascata): so eles voltam quando os dados descem
' (um rotulo apagado pelo usuario continua apagado); Stacked_SetLabels (ligar/desligar Rotulos) e "Alterar dados" limpam a lista.
Private Const BK_HDR As String = "LblPos"
Private Const BK_PREFIX As String = "tqq_ann_break_"
Private Const BK_LBL_OPT As String = "brklbl"
Private Const BK_TOP_ROOM As Double = 24        ' folga (pt) acima da area interna para o rotulo do total da coluna cortada
Private Const BK_RIGHT_MARGIN As Double = 10    ' barras: folga = largura do maior rotulo de total + esta margem
Private Const BK_TOP_SHARE As Double = 0.25     ' empilhadas: parte minima visivel do segmento de cima de uma coluna cortada (StackedFloor)
Public Const BREAK_ERR_LIMIT As String = "Corte de eixo: informe um valor maior que zero e menor que o maior valor do gráfico."
Private mBusy As Boolean                        ' Break_Set/Break_Remove em andamento: o refresh por evento (Calculate) nao interfere

Public Function Break_Limit(cht As Chart) As Double
    Break_Limit = Val(Registry_GetOption(cht, "vbreak", ""))
    If Break_Limit < 0 Then Break_Limit = 0
End Function

' Gancho de teste (Util_FailThinSyncForTest = "break-<ponto>:<chave>", como round-/bfly-): falha simulada num ponto do codigo para testar o
' caminho de erro. Dispara UMA vez (limpa o gancho antes de levantar o erro). Inerte fora do build Test.
Private Sub BreakTestHook(site As String, key As String)
    If BUILD_CONFIG = "Test" Then
        If Len(Util_FailThinSyncForTest) > 0 Then
            If Util_FailThinSyncForTest = "break-" & site & ":" & key Then
                Util_FailThinSyncForTest = ""
                Err.Raise vbObjectError + 9994, "ThinQQ", "Falha simulada (teste) " & site & " " & key
            End If
        End If
    End If
End Sub

' Celula do limite (BreakAt): a linha de meta da coluna LblPos (dentro do bloco: a limpeza a apaga junto e Block_Extend a ve como parte
' do bloco; vazio = sem corte, e MIN ignora a celula vazia). Nothing se a coluna ainda nao existe.
Private Function LimitCell(cht As Chart) As Range
    Dim blk As Range, c As Long
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    c = Block_FindCol(blk, BK_HDR)
    If c > 0 Then Set LimitCell = blk.Cells(1, c)
End Function

' Coluna LblPos (criada uma vez, a direita do bloco): MIN(TotalPos, $BreakAt), com BreakAt na linha de meta da propria coluna (LimitCell).
' Devolve a coluna no corpo. Uma gravacao so (cabecalho e formulas); transacional: uma falha depois de estender limpa a coluna e devolve o
' nome tqq_<id> a faixa anterior antes de o erro seguir.
Public Function Break_EnsureColumn(cht As Chart) As Long
    Dim blk As Range, body As Range, c As Long, n As Long, i As Long, lim As String, wb As Workbook, id As String, oldRef As String
    Dim f() As Variant, extended As Boolean, en As Long, es As String, ed As String
    Set blk = Registry_Block(cht)
    id = Registry_IdFromChart(cht)
    If blk Is Nothing Then Err.Raise vbObjectError + 6015, "Break_EnsureColumn", "Bloco inexistente: " & id
    c = Block_FindCol(blk, BK_HDR)
    If c > 0 Then Break_EnsureColumn = c: Exit Function
    Set wb = Registry_Workbook(cht)
    n = Block_Body(blk).Rows.Count - 1
    If n < 1 Then Err.Raise vbObjectError + 7002, "Break_EnsureColumn", "O gráfico não tem categorias."
    oldRef = wb.Names(ID_PREFIX & id).RefersTo
    On Error GoTo Fail
    Set blk = Block_Extend(wb, id, 1)
    extended = True
    Set body = Block_Body(blk)
    c = blk.Columns.Count
    lim = blk.Cells(1, c).Address(True, True)           ' meta da propria coluna: dentro do bloco
    ReDim f(1 To n + 1, 1 To 1)
    f(1, 1) = BK_HDR
    For i = 1 To n
        f(i + 1, 1) = "=MIN(" & body.Cells(i + 1, 3).Address(False, False) & "," & lim & ")"
    Next
    body.Cells(1, c).Resize(n + 1, 1).Formula = f
    BreakTestHook "ensure", "write"
    Break_EnsureColumn = c
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    If extended Then
        blk.Columns(c).Clear
        wb.Names(ID_PREFIX & id).RefersTo = oldRef
    End If
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' Topo positivo de cada categoria: soma dos valores > 0 nas empilhadas, maior valor entre as series nas agrupadas (series auxiliares fora).
Public Function Break_CategoryTops(cht As Chart) As Double()
    Dim tops() As Double, n As Long, i As Long, s As Series, v As Variant, x As Double, flat As Boolean
    flat = IsFlatKind(Registry_ChartType(cht))
    n = Geom_ChartCategoryCount(cht)
    If n < 1 Then ReDim tops(0 To 0): Break_CategoryTops = tops: Exit Function
    ReDim tops(1 To n)
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            v = s.Values
            For i = 1 To n
                x = ArrayValue(v, i)
                If x > 0 Then
                    If flat Then
                        If x > tops(i) Then tops(i) = x
                    Else
                        tops(i) = tops(i) + x
                    End If
                End If
            Next
        End If
    Next
    Break_CategoryTops = tops
End Function

Private Function IsFlatKind(kind As String) As Boolean
    IsFlatKind = (kind = "clustered" Or kind = "barclustered")
End Function

Private Function ArrayValue(v As Variant, i As Long) As Double
    If Not IsArray(v) Then Exit Function
    If LBound(v) + i - 1 > UBound(v) Then Exit Function
    ArrayValue = Util_ToDouble(v(LBound(v) + i - 1))
End Function

Private Function MaxTop(cht As Chart) As Double
    Dim tops() As Double, i As Long
    tops = Break_CategoryTops(cht)
    For i = LBound(tops) To UBound(tops)
        If tops(i) > MaxTop Then MaxTop = tops(i)
    Next
End Function

' Empilhadas: piso do corte em lim. Um segmento inteiro acima do corte some sem aviso (na demonstracao a coluna do cimento parecia so B2C):
' em cada categoria CORTADA (topo > lim) o segmento de cima (a ultima serie com valor > 0) precisa mostrar pelo menos BK_TOP_SHARE dele, ou
' seja, lim >= base dele + BK_TOP_SHARE x ele = topo - (1 - BK_TOP_SHARE) x ele. Devolve o maior piso entre as cortadas (cat = a categoria
' dele); 0 nas agrupadas (cada serie e uma coluna inteira) ou sem categoria cortada. Subir lim so tira categorias do conjunto: o piso em lim
' nunca cresce com lim (um valor >= piso(lim) continua valido).
Private Function StackedFloor(cht As Chart, lim As Double, ByRef cat As Long) As Double
    Dim n As Long, i As Long, s As Series, v As Variant, x As Double, tops() As Double, lastSeg() As Double, f As Double
    cat = 0
    If IsFlatKind(Registry_ChartType(cht)) Then Exit Function
    n = Geom_ChartCategoryCount(cht)
    If n < 1 Then Exit Function
    ReDim tops(1 To n): ReDim lastSeg(1 To n)
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            v = s.Values
            For i = 1 To n
                x = ArrayValue(v, i)
                If x > 0 Then tops(i) = tops(i) + x: lastSeg(i) = x
            Next
        End If
    Next
    For i = 1 To n
        If tops(i) > lim Then
            f = tops(i) - (1 - BK_TOP_SHARE) * lastSeg(i)
            If f > StackedFloor Then StackedFloor = f: cat = i
        End If
    Next
End Function

' Menor numero "redondo" >= v e < mx (2, 3, ... 6 algarismos significativos, arredondado para cima); v se nenhum couber.
Private Function NiceCeil(v As Double, mx As Double) As Double
    Dim d As Long, p As Double, r As Double
    NiceCeil = v
    If v <= 0 Then Exit Function
    For d = 2 To 6
        p = 10 ^ (Int(Log(v) / Log(10)) - d + 1)
        r = -Int(-v / p) * p
        If r < v Then r = r + p                      ' ruido de ponto flutuante
        If r < mx Then NiceCeil = r: Exit Function
    Next
End Function

' Numero no separador da localidade (Format$), sem a virgula final de um inteiro.
Private Function NumberText(v As Double) As String
    If v = Int(v) Then NumberText = Format$(v, "#,##0") Else NumberText = Format$(v, "#,##0.##")
End Function

' Texto da recusa do piso (7003): a coluna (barra) que define o piso e o menor limite aceito, ja arredondado para cima.
Private Function FloorText(cht As Chart, cat As Long, fl As Double) As String
    Dim nm As String, xv As Variant, what As String
    On Error Resume Next
    xv = cht.SeriesCollection(1).XValues
    If IsArray(xv) Then nm = CStr(xv(LBound(xv) + cat - 1))
    On Error GoTo 0
    If Len(nm) = 0 Then nm = CStr(cat)
    If Geom_ChartIsHorizontal(cht) Then what = "barra" Else what = "coluna"
    FloorText = "Corte de eixo: com esse valor o segmento de cima da " & what & " '" & nm & "' ficaria inteiro acima do corte e não apareceria. "
    FloorText = FloorText & "Informe pelo menos " & NumberText(NiceCeil(fl, MaxTop(cht))) & "."
End Function

' Sugestao de corte: 1,2 x o segundo maior topo (se menor que o maior), senao metade do maior; nas empilhadas, nunca abaixo do piso
' (StackedFloor: o segmento de cima da coluna cortada continua visivel); arredondada a 2 algarismos significativos (o InputBox mostra um
' numero redondo) quando o arredondado continua entre o piso e o maior topo, senao para cima (NiceCeil).
Public Function Break_Suggest(cht As Chart) As Double
    Dim tops() As Double, i As Long, mx As Double, second As Double, v As Double, p As Double, r As Double, fl As Double, cat As Long
    tops = Break_CategoryTops(cht)
    For i = LBound(tops) To UBound(tops)
        If tops(i) > mx Then
            second = mx: mx = tops(i)
        ElseIf tops(i) > second Then
            second = tops(i)
        End If
    Next
    If second > 0 And second * 1.2 < mx Then v = second * 1.2 Else v = mx / 2
    If v > 0 Then
        fl = StackedFloor(cht, v, cat)
        If fl > v Then v = fl
        p = 10 ^ (Int(Log(v) / Log(10)) - 1)
        r = Round(v / p) * p
        If r > 0 And r < mx And r >= StackedFloor(cht, r, cat) Then
            v = r
        ElseIf fl > 0 Then
            v = NiceCeil(v, mx)
        End If
    End If
    Break_Suggest = v
End Function

' Liga (ou muda) o corte em limit; limit <= 0 remove. Recusas: 7002 (tipo sem corte, matriz de capacidades) e 7003 (limite fora de
' 0 < limit < maior topo, ou, nas empilhadas, abaixo do piso de StackedFloor: o segmento de cima de uma coluna cortada sumiria; o aviso diz
' a coluna e o menor valor aceito), todas antes de gravar qualquer coisa. Recusar (e nao subir o limite sozinho) deixa o usuario decidir e
' nunca grava um valor que ele nao digitou. Transacional: uma falha depois de comecar volta a celula do limite, as opcoes (vbreak, vmin,
' vmax) e o eixo ao que eram antes de o erro seguir para o chamador.
Public Sub Break_Set(cht As Chart, limit As Double)
    Dim kind As String, ax As Axis, lc As Range, prevOpt As String, prevMin As String, prevMax As String, prevCell As Variant
    Dim prevAutoMax As Boolean, prevMaxVal As Double, prevAutoMin As Boolean, prevMinVal As Double, started As Boolean, wasBusy As Boolean
    Dim en As Long, es As String, ed As String, fl As Double, cat As Long
    kind = Registry_ChartType(cht)
    If Not Kinds_Supports(kind, "break") Then Err.Raise vbObjectError + 7002, "Break_Set", Kinds_NotApplicableText("break")
    If limit <= 0 Then Break_Remove cht: Exit Sub
    If limit >= MaxTop(cht) Then Err.Raise vbObjectError + 7003, "Break_Set", BREAK_ERR_LIMIT
    fl = StackedFloor(cht, limit, cat)
    If limit < fl - 0.000000001 * (1 + fl) Then Err.Raise vbObjectError + 7003, "Break_Set", FloorText(cht, cat, fl)
    wasBusy = mBusy: mBusy = True
    On Error GoTo Fail
    Break_EnsureColumn cht                                  ' transacional sozinha; fica criada (sem limite, LblPos = TotalPos)
    Set ax = cht.Axes(xlValue, xlPrimary)
    Set lc = LimitCell(cht)
    prevOpt = Registry_GetOption(cht, "vbreak", "")
    prevMin = Registry_GetOption(cht, "vmin", ""): prevMax = Registry_GetOption(cht, "vmax", "")
    prevCell = lc.Value
    prevAutoMax = ax.MaximumScaleIsAuto: prevMaxVal = ax.MaximumScale
    prevAutoMin = ax.MinimumScaleIsAuto: prevMinVal = ax.MinimumScale
    started = True
    lc.Value = limit
    Break_RepointTotals cht
    BreakTestHook "set", "cell"
    ' mesma escala e corte sao excludentes: a escala fixa (vmin/vmax) sai e o minimo volta ao automatico
    If Len(prevMin) > 0 Then Registry_SetOption cht, "vmin", ""
    If Len(prevMax) > 0 Then Registry_SetOption cht, "vmax", ""
    ax.MinimumScaleIsAuto = True
    ax.MaximumScale = limit
    BreakTestHook "set", "axis"
    Registry_SetOption cht, "vbreak", Deco_ScaleText(limit)
    On Error Resume Next                                    ' daqui em diante so layout (cada passo registra a propria falha)
    If Geom_ChartIsHorizontal(cht) Then Chart_SyncSecondaryX cht
    Chart_EnsureRoomAfterChange cht                         ' folga do total (Break_EnsureRoom) e as demais, para o eixo novo
    Break_SyncLabels cht
    Break_SyncMarks cht
    Labels_RefreshPositions cht
    Ann_Refresh cht
    If Err.Number <> 0 Then Errors_Log "Break_Set/layout" & vbTab & Err.Number & vbTab & Err.Description
    mBusy = wasBusy
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    If started Then
        If IsEmpty(prevCell) Then lc.ClearContents Else lc.Value = prevCell
        Registry_SetOption cht, "vbreak", prevOpt
        If Len(prevMin) > 0 Then Registry_SetOption cht, "vmin", prevMin
        If Len(prevMax) > 0 Then Registry_SetOption cht, "vmax", prevMax
        If prevAutoMin Then ax.MinimumScaleIsAuto = True Else ax.MinimumScale = prevMinVal
        If prevAutoMax Then ax.MaximumScaleIsAuto = True Else ax.MaximumScale = prevMaxVal
        If prevAutoMin Then ax.MinimumScaleIsAuto = True Else ax.MinimumScale = prevMinVal      ' de novo: a ordem pode ter sido recusada
        RelinkTotalsIfLost cht
        Break_SyncLabels cht
        Break_SyncMarks cht
        Ann_Refresh cht
        Err.Clear
    End If
    mBusy = wasBusy
    On Error GoTo 0
    Err.Raise en, es, ed
End Sub

' O total passa a seguir LblPos (vertical: Values; barras: XValues da auxiliar XY); o rotulo continua ligado ao Total (coluna 2). Sem limite,
' LblPos = TotalPos: apontar para ela nao muda nada. Medido: trocar Values/XValues de uma serie com rotulos LIGADOS apaga os rotulos dela E
' passa as ligacoes para rotulos de OUTRA serie (o rotulo de um segmento passou a mostrar o total de outra categoria). Por isso os rotulos do
' total saem antes (HasDataLabels = False) e sao religados depois (Stacked_SetTotals com a serie existente: rotulos, posicoes, negrito e a
' folga das barras). Ja apontando para LblPos nao mexe.
Private Sub Break_RepointTotals(cht As Chart)
    Dim s As Series, blk As Range, body As Range, c As Long, n As Long, ref As Range
    Set s = Chart_FindSeries(cht, "tqq_total")
    If s Is Nothing Then Exit Sub
    Set blk = Registry_Block(cht): Set body = Block_Body(blk)
    c = Block_FindCol(blk, BK_HDR)
    If c = 0 Then Exit Sub
    n = body.Rows.Count - 1
    Set ref = body.Cells(2, c).Resize(n, 1)
    If InStr(1, s.Formula, "!" & ref.Address(True, True) & ",", vbTextCompare) > 0 Then Exit Sub
    s.HasDataLabels = False
    If Kinds_IsHorizontal(Registry_ChartType(cht)) Then
        s.XValues = Block_Ref(ref)
    Else
        s.Values = Block_Ref(ref)
    End If
    BreakTestHook "repoint", "values"
    Stacked_SetTotals cht, True
End Sub

' Caminho de erro: a falha no meio de Break_RepointTotals (depois de tirar os rotulos do total) deixaria a serie sem rotulos; religa.
Private Sub RelinkTotalsIfLost(cht As Chart)
    Dim s As Series
    On Error Resume Next
    If Registry_GetOption(cht, "totals", "0") <> "1" Then Exit Sub
    Set s = Chart_FindSeries(cht, "tqq_total")
    If s Is Nothing Then Exit Sub
    If Not s.HasDataLabels Then Stacked_SetTotals cht, True
    If Err.Number <> 0 Then Errors_Log "Break_Set/desfazer totais" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Remove o corte: eixo de volta ao automatico (antes de tudo: se falhar, nada mudou), limite apagado, opcao limpa, marcas apagadas, rotulos
' ocultos pelo corte de volta e a folga de cima devolvida (layout automatico; as demais folgas sao refeitas). Sem corte (nem opcao, nem limite,
' nem marca, nem rotulo oculto) nao faz nada: o eixo e o layout de um grafico sem corte ficam como estao.
Public Sub Break_Remove(cht As Chart)
    Dim i As Long, lc As Range, axisOurs As Boolean, wasBusy As Boolean, en As Long, es As String, ed As String
    Set lc = LimitCell(cht)
    axisOurs = (Len(Registry_GetOption(cht, "vbreak", "")) > 0)
    If Not lc Is Nothing Then
        If Not IsEmpty(lc.Value) Then axisOurs = True
    End If
    If Not axisOurs Then
        If Not HasAnyMark(cht) And Len(Registry_GetOption(cht, BK_LBL_OPT, "")) = 0 Then Exit Sub
    End If
    wasBusy = mBusy: mBusy = True
    On Error GoTo Fail
    If axisOurs Then cht.Axes(xlValue, xlPrimary).MaximumScaleIsAuto = True
    If Not lc Is Nothing Then lc.ClearContents
    Registry_SetOption cht, "vbreak", ""
    On Error Resume Next
    For i = cht.Shapes.Count To 1 Step -1
        If Left$(cht.Shapes(i).Name, Len(BK_PREFIX)) = BK_PREFIX Then cht.Shapes(i).Delete
    Next
    Break_SyncLabels cht                             ' sem limite: os rotulos ocultos pelo corte voltam e os da base voltam a posicao normal
    If axisOurs Then
        cht.PlotArea.Position = xlChartElementPositionAutomatic     ' a folga de cima (so reduz) volta; as demais folgas sao refeitas
        If Geom_ChartIsHorizontal(cht) Then Chart_SyncSecondaryX cht
        Chart_EnsureRoomAfterChange cht
    End If
    Labels_RefreshPositions cht
    Ann_Refresh cht
    If Err.Number <> 0 Then Errors_Log "Break_Remove" & vbTab & Err.Number & vbTab & Err.Description
    mBusy = wasBusy
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    mBusy = wasBusy
    Err.Raise en, es, ed
End Sub

' Reaplica o maximo fixo se o eixo voltou ao automatico (ex.: o usuario mexeu no eixo) ou difere do limite. Sem corte nao faz nada.
Public Sub Break_ApplyAxis(cht As Chart)
    Dim lim As Double, ax As Axis
    lim = Break_Limit(cht)
    If lim <= 0 Then Exit Sub
    Set ax = cht.Axes(xlValue, xlPrimary)
    If ax.MaximumScaleIsAuto Or Abs(ax.MaximumScale - lim) > 0.000000001 * (1 + lim) Then
        ax.MaximumScale = lim
        If Geom_ChartIsHorizontal(cht) Then Chart_SyncSecondaryX cht
    End If
End Sub

' Folga para o rotulo do total da coluna cortada (so com a auxiliar tqq_total: agrupadas e totais desligados nao tem rotulo acima do corte):
' 24 pt acima da area interna (colunas) ou a largura do maior rotulo de total + margem a direita (barras). So aumenta a folga. Chamada por
' Chart_EnsureRoomAfterChange e pelo refresh (ligar os totais depois do corte). Nunca propaga erro.
Public Sub Break_EnsureRoom(cht As Chart)
    Dim k As Long, delta As Double, need As Double
    On Error Resume Next
    If Break_Limit(cht) <= 0 Then Exit Sub
    If Chart_FindSeries(cht, "tqq_total") Is Nothing Then Exit Sub
    If Geom_ChartIsHorizontal(cht) Then
        need = Labels_MaxTotalWidth(cht)
        If need <= 0 Then Exit Sub
        need = need + BK_RIGHT_MARGIN
        For k = 1 To 3
            delta = need - (Geom_ChartRightEdge(cht) - (cht.PlotArea.InsideLeft + cht.PlotArea.InsideWidth))
            If delta <= 0.5 Then Exit For
            If Not Deco_ShrinkPlotWidth(cht, delta) Then Exit For
        Next
    Else
        For k = 1 To 3
            delta = BK_TOP_ROOM - cht.PlotArea.InsideTop
            If delta <= 0.01 Then Exit For
            cht.PlotArea.Height = cht.PlotArea.Height - delta
            cht.PlotArea.Top = cht.PlotArea.Top + delta
        Next
    End If
    If Err.Number <> 0 Then Errors_Log "Break_EnsureRoom" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' ---------- rotulos dos segmentos ----------

Private Function ListHas(list As String, key As String) As Boolean
    ListHas = InStr("," & list & ",", "," & key & ",") > 0
End Function

Private Function ListAdd(list As String, key As String) As String
    If ListHas(list, key) Then ListAdd = list: Exit Function
    If Len(list) = 0 Then ListAdd = key Else ListAdd = list & "," & key
End Function

Private Function ListRemove(list As String, key As String) As String
    Dim s As String
    s = Replace("," & list & ",", "," & key & ",", ",")
    If Len(s) <= 1 Then Exit Function
    ListRemove = Mid$(s, 2, Len(s) - 2)
End Function

' Cor da fonte do rotulo do segmento na posicao pos: dentro da coluna (Center, InsideBase) contraste com o preenchimento; fora (OutsideEnd,
' agrupadas) cinza-escuro.
Private Function LabelFontColor(s As Series, pos As Long) As Long
    If pos = xlLabelPositionOutsideEnd Then
        LabelFontColor = RGB(64, 64, 64)
    Else
        LabelFontColor = Style_ContrastFont(CLng(s.Format.Fill.ForeColor.RGB))
    End If
End Function

' Religa o rotulo de um ponto oculto pelo corte como Stacked_SetLabels o faria (um ponto religado nao herda o conteudo da serie): ligado a
' Rnd_j com a soma consistente, senao o valor nativo; formato dos segmentos, fonte 9 e a cor da posicao. Se a formatacao falha depois de
' religar, o rotulo volta a ficar oculto e o erro segue (o par continua na lista).
Private Sub RestorePointLabel(cht As Chart, s As Series, j As Long, i As Long, pos As Long, rcRef As String)
    Dim en As Long, es As String, ed As String
    s.Points(i).HasDataLabel = True
    On Error GoTo Fail
    BreakTestHook "restore", j & "." & i
    With s.Points(i).DataLabel
        If Len(rcRef) > 0 Then
            .Formula = rcRef
        Else
            .ShowValue = True: .ShowSeriesName = False: .ShowCategoryName = False
        End If
        Style_AssignNumberFormat s.Points(i).DataLabel, Labels_SegmentFormat(cht)
        .Font.Size = 9
        .Font.Color = LabelFontColor(s, pos)
        .Position = pos
    End With
    Exit Sub
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    s.Points(i).HasDataLabel = False
    On Error GoTo 0
    Err.Raise en, es, ed
End Sub

' Rotulos nativos dos segmentos (opcao labels ligada): segmento com a base >= limite fica sem rotulo; segmento atravessado pelo limite vai
' para a base (InsideBase, fonte por contraste); os demais ficam na posicao normal (Center nas empilhadas, OutsideEnd nas agrupadas). Sem
' limite, todos voltam ao normal. So o ThinQQ religa o que ELE ocultou (lista brklbl): um rotulo ausente fora da lista foi apagado pelo
' usuario e fica apagado. Transacional por ponto: um ponto cuja leitura falha fica como esta; um par so entra na lista quando ocultar deu
' certo e so sai quando religar deu certo; a lista e gravada tambem na saida por erro. So escreve o que mudou. Nunca propaga erro.
Public Sub Break_SyncLabels(cht As Chart)
    Dim lim As Double, flat As Boolean, n As Long, i As Long, j As Long, s As Series, v As Variant, x As Double, base() As Double
    Dim want As Long, hide As Boolean, has As Boolean, readOk As Boolean, normalPos As Long, cur As Long, fc As Long
    Dim list0 As String, list As String, key As String, rs As Boolean, blk As Range, body As Range, rc As Long, rcRef As String, sheetRef As String
    Dim en As Long, ed As String
    On Error GoTo Done
    list0 = Registry_GetOption(cht, BK_LBL_OPT, ""): list = list0
    If Registry_GetOption(cht, "labels", "1") <> "1" Then GoTo Done
    flat = IsFlatKind(Registry_ChartType(cht))
    normalPos = IIf(flat, xlLabelPositionOutsideEnd, xlLabelPositionCenter)
    lim = Break_Limit(cht)
    n = Geom_ChartCategoryCount(cht)
    If n < 1 Then GoTo Done
    ' soma consistente ligada e coerente (como em Stacked_SetLabels): o rotulo religado liga a Rnd_j
    If Registry_GetOption(cht, "roundsum", "0") = "1" Then
        Set blk = Registry_Block(cht)
        If Not blk Is Nothing Then
            If Round_ColumnCount(blk) = Round_SeriesCount(cht) Then
                rs = True
                Set body = Block_Body(blk)
                sheetRef = "='" & Replace(blk.Worksheet.Name, "'", "''") & "'!"
            End If
        End If
    End If
    ReDim base(1 To n)
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            j = j + 1
            v = Empty
            On Error Resume Next
            v = s.Values
            Err.Clear
            On Error GoTo Done
            rc = 0
            If rs Then rc = Round_RCol(blk, j)
            For i = 1 To n
                x = ArrayValue(v, i)
                hide = False: want = normalPos: has = False: readOk = False: cur = 0
                If lim > 0 And x > 0 Then
                    If base(i) >= lim Then
                        hide = True
                    ElseIf base(i) + x > lim Then
                        want = xlLabelPositionInsideBase
                    End If
                End If
                If x > 0 And Not flat Then base(i) = base(i) + x
                key = j & "." & i
                On Error Resume Next
                has = s.Points(i).HasDataLabel
                readOk = (Err.Number = 0)
                Err.Clear
                If readOk Then
                    If hide Then
                        If has Then
                            s.Points(i).HasDataLabel = False
                            If Err.Number = 0 Then list = ListAdd(list, key)
                        End If
                    ElseIf has Then
                        cur = s.Points(i).DataLabel.Position
                        If Err.Number = 0 Then
                            If cur <> want Then
                                s.Points(i).DataLabel.Position = want
                                s.Points(i).DataLabel.Font.Color = LabelFontColor(s, want)
                            ElseIf want = xlLabelPositionInsideBase Then
                                fc = LabelFontColor(s, want)       ' a paleta pode ter mudado o preenchimento (agrupadas: a serie nao recolore a base)
                                If s.Points(i).DataLabel.Font.Color <> fc Then s.Points(i).DataLabel.Font.Color = fc
                            End If
                        End If
                    ElseIf ListHas(list, key) Then
                        rcRef = ""
                        If rc > 0 Then rcRef = sheetRef & body.Cells(i + 1, rc).Address(True, True)
                        RestorePointLabel cht, s, j, i, want, rcRef
                        If Err.Number = 0 Then list = ListRemove(list, key)
                    End If
                End If
                If Err.Number <> 0 Then Errors_Log "Break_SyncLabels" & vbTab & key & vbTab & Err.Number & vbTab & Err.Description
                Err.Clear
                On Error GoTo Done
            Next
        End If
    Next
Done:
    If Err.Number <> 0 Then en = Err.Number: ed = Err.Description
    On Error Resume Next
    If list <> list0 Then Registry_SetOption cht, BK_LBL_OPT, list
    If en <> 0 Then Errors_Log "Break_SyncLabels" & vbTab & en & vbTab & ed
End Sub

' Pontos (",3,5,") da serie de dados j (ordem das series nao auxiliares) inteiros acima do corte: a base do segmento (soma dos positivos das
' series anteriores, empilhadas) >= limite. Agrupadas e sem corte: "". Stacked_SetLabels nao liga esses rotulos (Labels_LinkPoints skipPoints).
Public Function Break_CutPointList(cht As Chart, j As Long) As String
    Dim lim As Double, n As Long, i As Long, s As Series, k As Long, base() As Double, x As Double, v As Variant, out As String
    On Error GoTo Fail
    lim = Break_Limit(cht)
    If lim <= 0 Or IsFlatKind(Registry_ChartType(cht)) Then Exit Function
    n = Geom_ChartCategoryCount(cht)
    If n < 1 Then Exit Function
    ReDim base(1 To n)
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            k = k + 1
            v = s.Values
            For i = 1 To n
                x = ArrayValue(v, i)
                If k = j Then
                    If x > 0 And base(i) >= lim Then out = out & i & ","
                ElseIf x > 0 Then
                    base(i) = base(i) + x
                End If
            Next
            If k = j Then Exit For
        End If
    Next
    If Len(out) > 0 Then Break_CutPointList = "," & out
    Exit Function
Fail:
    Break_CutPointList = ""
End Function

' ---------- marcas ----------

Private Function HasAnyMark(cht As Chart) As Boolean
    Dim shp As Shape
    For Each shp In cht.Shapes
        If Left$(shp.Name, Len(BK_PREFIX)) = BK_PREFIX Then HasAnyMark = True: Exit Function
    Next
End Function

' Categoria de uma forma de marca (pelo nome tqq_ann_break_<i>[_w|_b]); 0 se o nome nao tem essa forma.
Private Function MarkCategory(nm As String) As Long
    Dim rest As String
    If Left$(nm, Len(BK_PREFIX)) <> BK_PREFIX Then Exit Function
    rest = Mid$(nm, Len(BK_PREFIX) + 1)
    If Right$(rest, 2) = "_w" Or Right$(rest, 2) = "_b" Then rest = Left$(rest, Len(rest) - 2)
    If Len(rest) = 0 Or Not rest Like String$(Len(rest), "#") Then Exit Function
    MarkCategory = CLng(rest)
End Function

' Uma marca por categoria com topo acima do limite: tqq_ann_break_<i>_w (branca, 3 pt, criada primeiro: fica atras), tqq_ann_break_<i> (cinza)
' e tqq_ann_break_<i>_b (cinza). Apaga as que sobram (categoria que desceu, que deixou de existir, ou marca incompleta: apagada e recriada) e
' cria as que faltam; Ann_Refresh posiciona (Ann_LayoutBreak). Nunca propaga erro.
Public Sub Break_SyncMarks(cht As Chart)
    Dim lim As Double, tops() As Double, i As Long, k As Long, nm As String, shp As Shape, sfx As Variant, parts() As Long, n As Long, cat As Long
    On Error GoTo Fail
    lim = Break_Limit(cht)
    tops = Break_CategoryTops(cht)
    n = UBound(tops)
    If n < 0 Then n = 0
    ReDim parts(0 To n)
    For k = cht.Shapes.Count To 1 Step -1
        nm = cht.Shapes(k).Name
        If Left$(nm, Len(BK_PREFIX)) = BK_PREFIX Then
            cat = MarkCategory(nm)
            If cat < 1 Or cat > n Or lim <= 0 Then
                cht.Shapes(k).Delete
            ElseIf tops(cat) <= lim Then
                cht.Shapes(k).Delete
            Else
                parts(cat) = parts(cat) + 1
            End If
        End If
    Next
    If lim <= 0 Then Exit Sub
    For i = 1 To n
        If tops(i) > lim And parts(i) <> 3 Then
            nm = BK_PREFIX & i
            For k = cht.Shapes.Count To 1 Step -1
                If MarkCategory(cht.Shapes(k).Name) = i Then cht.Shapes(k).Delete
            Next
            For Each sfx In Array("_w", "", "_b")
                Set shp = cht.Shapes.AddLine(0, 0, 10, 0)
                shp.Name = nm & sfx
                shp.AlternativeText = "kind=break;cat=" & i & ";part=" & IIf(sfx = "_w", "w", IIf(sfx = "_b", "b", "a"))
                If sfx = "_w" Then
                    shp.Line.ForeColor.RGB = RGB(255, 255, 255): shp.Line.Weight = 3
                Else
                    shp.Line.ForeColor.RGB = RGB(89, 89, 89): shp.Line.Weight = 1
                End If
            Next
        End If
    Next
    Exit Sub
Fail:
    Errors_Log "Break_SyncMarks" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Extensao da marca da categoria cat no sentido das categorias: centro (coord) e meia largura (half). Empilhadas: a coluna inteira. Agrupadas:
' do primeiro ao ultimo ponto CORTADO da categoria (a marca nao atravessa o espaco acima de uma coluna vizinha mais baixa), com a geometria
' do Excel para o grupo: largura de cada coluna w = fatia / (S - (S - 1) x sobreposicao + espaco), passo w x (1 - sobreposicao) (medido).
' Barras: o eixo de categorias invertido (categoria 1 no topo) inverte tambem as series dentro da categoria: a serie 1 fica em cima (medido),
' na mesma ordem das colunas (da esquerda para a direita = de cima para baixo). False se a categoria nao tem ponto cortado (agrupadas).
Public Function Break_MarkExtent(cht As Chart, cat As Long, ByRef coord As Double, ByRef half As Double) As Boolean
    Dim lim As Double, s As Series, j As Long, ns As Long, first As Long, last As Long, w As Double, stp As Double, slot As Double
    Dim ov As Double, gap As Double, c0 As Double, lo As Double, hi As Double, n As Long
    coord = Geom_ChartCatCoord(cht, cat)
    half = Geom_ChartBarHalfWidth(cht)
    If Not IsFlatKind(Registry_ChartType(cht)) Then Break_MarkExtent = True: Exit Function
    lim = Break_Limit(cht)
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            ns = ns + 1
            If ArrayValue(s.Values, cat) > lim Then
                If first = 0 Then first = ns
                last = ns
            End If
        End If
    Next
    If first = 0 Or ns < 1 Then Exit Function
    n = Geom_ChartCategoryCount(cht)
    If Geom_ChartIsHorizontal(cht) Then slot = cht.PlotArea.InsideHeight / n Else slot = cht.PlotArea.InsideWidth / n
    gap = cht.ChartGroups(1).GapWidth / 100
    ov = cht.ChartGroups(1).Overlap / 100
    w = slot / (ns - (ns - 1) * ov + gap)
    stp = w * (1 - ov)
    c0 = coord - (w + (ns - 1) * stp) / 2                 ' borda inicial do grupo (esquerda nas colunas, de cima nas barras)
    lo = c0 + (first - 1) * stp
    hi = c0 + (last - 1) * stp + w
    coord = (lo + hi) / 2
    half = (hi - lo) / 2
    Break_MarkExtent = True
End Function

' ---------- refresh ----------

' O que no corte depende dos DADOS e muda depois de inserir (Core_SafeRefresh): o maximo fixo (se o eixo voltou ao automatico), a folga do
' total, os rotulos dos segmentos e as marcas. So com a coluna LblPos criada (um grafico novo de "Alterar dados" recebe o corte por
' Break_TryReapply). Nunca propaga erro; guarda de reentrada (as marcas/rotulos disparam Calculate) e nada no meio de Labels_LinkPoints (cada
' rotulo ligado dispara o refresh: mexer nos rotulos ali apagaria o que esta sendo ligado e multiplicaria o custo; quem liga sincroniza no fim).
Public Sub Break_RefreshAfterDataChange(cht As Chart)
    If mBusy Then Exit Sub
    If Labels_IsLinking() Then Exit Sub
    mBusy = True
    On Error Resume Next
    If Kinds_Supports(Registry_ChartType(cht), "break") Then
        If Break_Limit(cht) > 0 Then
            If Not LimitCell(cht) Is Nothing Then
                Break_ApplyAxis cht
                Break_EnsureRoom cht
                Break_SyncLabels cht
                Break_SyncMarks cht
            End If
        ElseIf Len(Registry_GetOption(cht, BK_LBL_OPT, "")) > 0 Or HasAnyMark(cht) Then
            Break_SyncLabels cht                       ' corte removido sem o refresh final: os rotulos ocultos voltam
            Break_SyncMarks cht
        End If
    End If
    If Err.Number <> 0 Then Errors_Log "Break_RefreshAfterDataChange" & vbTab & Err.Number & vbTab & Err.Description
    mBusy = False
End Sub

' Depois de "Alterar dados" o grafico novo tem a opcao copiada mas nenhuma coluna/marca: reaplica; se o limite ja nao vale para os dados
' novos (7003) ou a aplicacao falha, limpa a opcao e registra. Nunca propaga erro.
Public Sub Break_TryReapply(cht As Chart)
    Dim lim As Double
    On Error GoTo Fail
    lim = Break_Limit(cht)
    If lim <= 0 Then Exit Sub
    If Not Kinds_Supports(Registry_ChartType(cht), "break") Then Registry_SetOption cht, "vbreak", "": Exit Sub
    Break_Set cht, lim
    Exit Sub
Fail:
    Errors_Log "Break_TryReapply" & vbTab & Err.Number & vbTab & Err.Description
    On Error Resume Next
    Registry_SetOption cht, "vbreak", ""
End Sub
