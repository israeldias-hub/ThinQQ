Attribute VB_Name = "modChartArea"
Option Explicit
' Area empilhada e area 100% (kinds area, area100).
' Bloco (corpo): linha 1 = cabecalhos; colunas (S = numero de series de dados):
'  1 Categoria | 2 Total (=AGGREGATE(9,6,linha): ignora erros e texto) | 3 TotalPos (posicao plotada do rotulo do total:
'  o topo VISIVEL da pilha = o maior dos acumulados SUM(Val_1..Val_j), que e o proprio Total quando nao ha valor negativo;
'  se nenhum acumulado e positivo, o menor deles; 1 em area100) | 4..3+S Val_j (valor da serie j; area100: participacao = valor / Total, 0 se o Total e 0
'  ou a celula nao e numero) | 4+S..3+2S Mid_j (meio da faixa j na pilha: SUM(Val_1..Val_j) - Val_j/2, onde fica o rotulo)
'  | 4+2S Pos | 5+2S VLine (media do Total; #N/D em area100).
' Os rotulos de segmento sao series auxiliares de linha invisivel (tqq_lbl_j) em Mid_j com o rotulo ligado a Val_j; o total
' e a auxiliar tqq_total do Stacked (mesmas colunas 1..3).
' Como o Excel empilha AREAS com valor negativo (medido no Excel 2019 amostrando os pixels do grafico exportado): soma
' algebricamente, sem separar positivos e negativos como as colunas. A faixa j e o poligono entre os acumulados c_(j-1) e c_j
' e as series seguintes pintam POR CIMA das anteriores (com 10, -8 e 5 os acumulados sao 10, 2 e 7: a faixa de -8 fica entre
' 2 e 10 e a de 5 entre 2 e 7 cobre a parte baixa dela; o que se ve acima de 7 e so a faixa de -8, ate 10). Logo o topo visivel
' da pilha e o MAIOR acumulado (10), nao o Total (7): o rotulo do total e as setas ancoram nele (Ann_GetCategoryData), enquanto o
' texto do rotulo continua sendo o Total. O Mid_j e o meio algebrico da faixa (c_(j-1) + c_j) / 2: com valor negativo ele pode cair
' sobre uma faixa posterior (limite do proprio grafico empilhado do Excel). Celula com texto ou #N/D: o Excel desenha a faixa
' com espessura 0 naquele ponto (nao interpola; medido) e o bloco usa Val = 0, igual.

Private Const AREA_THIN_OPT As String = "thinlbl"   ' pares j.i dos rotulos de ponta que o ThinQQ ocultou (Area_SyncEndLabels)
Private Const AREA_SIDE_OFFSET As Double = 6        ' o rotulo Right/Left comeca ~5,5 pt depois do ponto (medido nos totais XY)
Private Const AREA_TEXT_HEIGHT As Double = 12       ' altura do texto de um rotulo de 9 pt: centros mais proximos que isso se sobrepoem

Public Function Area_BuildBlock(wb As Workbook, id As String, data As clsChartData, kind As String) As Range
    Dim blk As Range, body As Range, i As Long, j As Long, n As Long, ns As Long, is100 As Boolean
    Dim rowCells As Range, srcAddr As String, totAddr As String, totRange As String, valAddr As String, cumExpr As String
    Dim midC As Long, posC As Long, vlC As Long
    is100 = (kind = "area100")
    n = data.CategoryCount
    ns = data.SeriesCount
    Set blk = CalcSheet_AllocateBlock(wb, id, kind, Mid$(ChartData_ExternalAddress(data.Source), 2), n + 1, 5 + 2 * ns)
    Set body = Block_Body(blk)
    body.Cells(1, 1).Value = "Categoria": body.Cells(1, 2).Value = "Total": body.Cells(1, 3).Value = "TotalPos"
    For j = 1 To ns: body.Cells(1, Area_ValCol(blk, j)).Value = "Val" & j: Next
    ' colunas lidas UMA vez (cada Area_*Col percorre os cabecalhos): Mid_j = midC + j - 1
    midC = Area_MidCol(blk, 1): posC = Area_PosCol(blk): vlC = Area_VLineCol(blk)
    For j = 1 To ns: body.Cells(1, midC + j - 1).Value = "Mid" & j: Next
    body.Cells(1, posC).Value = "Pos": body.Cells(1, vlC).Value = "VLine"
    totRange = body.Cells(2, 2).Resize(n, 1).Address(True, True)       ' mesma planilha: sem nome de planilha
    For i = 1 To n
        body.Cells(i + 1, 1).Formula = ChartData_ExternalAddress(data.Source.Cells(i + 1, 1))
        Set rowCells = data.Source.Cells(i + 1, 2).Resize(1, ns)
        body.Cells(i + 1, 2).Formula = "=AGGREGATE(9,6," & Mid$(ChartData_ExternalAddress(rowCells), 2) & ")"
        totAddr = body.Cells(i + 1, 2).Address(False, False)
        If is100 Then
            body.Cells(i + 1, 3).Formula = "=1"
        Else
            ' maior acumulado (se positivo) ou o menor: MAX/MIN sobre Mid_j + Val_j / 2 (= o acumulado c_j), em SUMPRODUCT para
            ' avaliar a matriz sem Ctrl+Shift+Enter
            cumExpr = body.Cells(i + 1, midC).Resize(1, ns).Address(False, False) & "+" & body.Cells(i + 1, Area_ValCol(blk, 1)).Resize(1, ns).Address(False, False) & "/2"
            body.Cells(i + 1, 3).Formula = "=IF(SUMPRODUCT(MAX(" & cumExpr & "))>0,SUMPRODUCT(MAX(" & cumExpr & ")),SUMPRODUCT(MIN(" & cumExpr & ")))"
        End If
        For j = 1 To ns
            srcAddr = Mid$(ChartData_ExternalAddress(data.Source.Cells(i + 1, 1 + j)), 2)
            If is100 Then
                body.Cells(i + 1, Area_ValCol(blk, j)).Formula = "=IF(AND(ISNUMBER(" & srcAddr & ")," & totAddr & "<>0)," & srcAddr & "/" & totAddr & ",0)"
            Else
                body.Cells(i + 1, Area_ValCol(blk, j)).Formula = "=IF(ISNUMBER(" & srcAddr & ")," & srcAddr & ",0)"
            End If
            valAddr = body.Cells(i + 1, Area_ValCol(blk, j)).Address(False, False)
            body.Cells(i + 1, midC + j - 1).Formula = "=SUM(" & body.Cells(i + 1, Area_ValCol(blk, 1)).Address(False, False) & ":" & valAddr & ")-" & valAddr & "/2"
        Next
        body.Cells(i + 1, posC).Formula = "=" & (n - i) & ".5"      ' independe do separador decimal da localidade
        If is100 Then
            body.Cells(i + 1, vlC).Formula = "=NA()"
        Else
            body.Cells(i + 1, vlC).Formula = "=AVERAGE(" & totRange & ")"
        End If
    Next
    Set Area_BuildBlock = blk
End Function

' Numero de series de dados: cabecalhos Val1, Val2, ... contiguos a partir da coluna 4 (a Fase 3 estende o bloco a direita; a largura
' ja nao diz quantas series ha). Reserva (bloco sem cabecalho Val1): (Columns.Count - 5) \ 2, como antes.
Public Function Area_SeriesCount(blk As Range) As Long
    Dim j As Long, body As Range
    Set body = Block_Body(blk)
    For j = 4 To body.Columns.Count
        If CStr(body.Cells(1, j).Value) <> "Val" & (j - 3) Then Exit For
        Area_SeriesCount = Area_SeriesCount + 1
    Next
    If Area_SeriesCount = 0 Then Area_SeriesCount = (blk.Columns.Count - 5) \ 2
End Function

Public Function Area_ValCol(blk As Range, j As Long) As Long
    Area_ValCol = 3 + j
End Function

Public Function Area_MidCol(blk As Range, j As Long) As Long
    Area_MidCol = 3 + Area_SeriesCount(blk) + j
End Function

Public Function Area_PosCol(blk As Range) As Long
    Area_PosCol = Block_FindCol(blk, "Pos")
    If Area_PosCol = 0 Then Area_PosCol = blk.Columns.Count - 1
End Function

Public Function Area_VLineCol(blk As Range) As Long
    Area_VLineCol = Block_FindCol(blk, "VLine")
    If Area_VLineCol = 0 Then Area_VLineCol = blk.Columns.Count
End Function

Public Function Area_Insert(ws As Worksheet, data As clsChartData, kind As String) As ChartObject
    Dim id As String, co As ChartObject, cht As Chart, s As Series, j As Long, anchor As Range
    Dim en As Long, es As String, ed As String
    If Kinds_Family(kind) <> "area" Then Err.Raise 5, "Area_Insert", "Tipo de area desconhecido: " & kind        ' antes de alocar o bloco
    On Error GoTo Fail
    id = Registry_NewId()
    Area_BuildBlock ws.Parent, id, data, kind
    Util_TestHookAfterBlock
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, Chart_FreeTop(ws, anchor.Left, anchor.Top, 460, 280), 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    For j = 1 To data.SeriesCount
        Set s = cht.SeriesCollection.NewSeries
        s.Name = ChartData_ExternalAddress(data.SeriesNameCell(j))
        s.Values = ChartData_ExternalAddress(data.SeriesValueCells(j))
        s.XValues = ChartData_ExternalAddress(data.CategoryCells())
    Next
    cht.ChartType = IIf(kind = "area100", xlAreaStacked100, xlAreaStacked)
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlZero
    Style_ApplyBase cht
    Util_TestHookAfterChart
    ' Fixa o eixo "entre categorias" (o padrao do Excel para area criada por codigo, mas o geometria das setas e dos rotulos
    ' le o valor do eixo via Geom_ChartBetween): a faixa comeca a meia fatia da borda, o que da lugar aos rotulos das pontas e
    ' a seta de diferenca, como nas colunas.
    cht.Axes(xlCategory).AxisBetweenCategories = True
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionRight
    cht.Legend.Font.Size = 9
    Registry_SetOption cht, "numfmt", "int"
    Registry_SetOption cht, "palette", "padrao"
    Registry_SetOption cht, "n", CStr(data.CategoryCount)
    Registry_SetOption cht, "legend", "right"
    Style_ApplySeriesColors cht, "padrao"
    Area_SetLabels cht, True
    Area_SetTotals cht, True
    If Kinds_Supports(kind, "ind100") Then Deco_SetInd100 cht, True        ' indicador 100%: ligado por padrao na area 100%
    Set Area_Insert = co
    Exit Function
Fail:
    en = Err.Number: es = Err.Source: ed = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Len(id) > 0 Then CalcSheet_DeleteBlock ws.Parent, id
    On Error GoTo 0
    Err.Raise en, es, ed
End Function

' Cores da area: paleta (Style_ApplySeriesColors: preenchimento e borda branca) e, depois, a cor da fonte de cada rotulo de
' segmento por contraste com a cor nova da faixa.
Public Sub Area_ApplyColors(cht As Chart, paletteName As String)
    Dim j As Long, ls As Series
    Style_ApplySeriesColors cht, paletteName
    For j = 1 To Area_SeriesCount(Registry_Block(cht))
        Set ls = Chart_FindSeries(cht, "tqq_lbl_" & j)
        If Not ls Is Nothing Then
            If ls.HasDataLabels Then ls.DataLabels.Font.Color = Style_ContrastFont(Style_DataSeriesColor(cht, j))
        End If
    Next
End Sub

' Rotulos de segmento: uma auxiliar de linha invisivel tqq_lbl_j por serie, com os pontos em Mid_j (meio da faixa) e o
' rotulo ligado a celula Val_j (valor; em area100, a participacao), posicao central (nas pontas, para dentro da faixa),
' fonte por contraste com a faixa.
' O formato e o dos segmentos (Labels_SegmentFormat: numfmt; em area100, pct0/pct1). Desligar apaga as auxiliares.
Public Sub Area_SetLabels(cht As Chart, turnOn As Boolean)
    Dim s As Series, blk As Range, body As Range, n As Long, ns As Long, j As Long, i As Long, fmt As String, rc As Long
    If Not turnOn Then
        For i = cht.SeriesCollection.Count To 1 Step -1
            Set s = cht.SeriesCollection(i)
            If Left$(s.Name, 8) = "tqq_lbl_" Then s.Delete
        Next
        If Len(Registry_GetOption(cht, AREA_THIN_OPT, "")) > 0 Then Registry_SetOption cht, AREA_THIN_OPT, ""
        Registry_SetOption cht, "labels", "0"
        Exit Sub
    End If
    Set blk = Registry_Block(cht)
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    ns = Area_SeriesCount(blk)
    fmt = Labels_SegmentFormat(cht)
    For j = 1 To ns
        Set s = Chart_FindSeries(cht, "tqq_lbl_" & j)
        If s Is Nothing Then Set s = Chart_AddHelperSeries(cht, "tqq_lbl_" & j, body.Cells(2, Area_MidCol(blk, j)).Resize(n, 1))
        rc = 0: If Registry_GetOption(cht, "roundsum", "0") = "1" Then rc = Round_RCol(blk, j)        ' soma consistente: Rnd_j (maior resto, modRounding)
        Labels_LinkPoints s, body.Cells(2, IIf(rc > 0, rc, Area_ValCol(blk, j))).Resize(n, 1), fmt, xlLabelPositionCenter, Style_ContrastFont(Style_DataSeriesColor(cht, j))
        ' A faixa começa no PONTO da primeira categoria e termina no da última: um rótulo centrado ali ficaria meio fora da
        ' faixa (texto claro sobre o fundo branco some). Nas pontas o rótulo fica para dentro: à direita do primeiro ponto e
        ' à esquerda do último.
        If n >= 2 Then
            s.Points(1).DataLabel.Position = xlLabelPositionRight
            s.Points(n).DataLabel.Position = xlLabelPositionLeft
        End If
    Next
    Chart_HideHelperLegendEntries cht
    Registry_SetOption cht, "labels", "1"
    If Len(Registry_GetOption(cht, AREA_THIN_OPT, "")) > 0 Then Registry_SetOption cht, AREA_THIN_OPT, ""     ' todos religados acima
    Area_SyncEndLabels cht                           ' pontas que encostam no vizinho: ocultas (e na lista)
End Sub

' Gancho de teste (Util_FailThinSyncForTest = "area-<ponto>:<chave>", como break-/round-): falha simulada num ponto de Area_SyncEndLabels.
' Dispara UMA vez (limpa o gancho antes de levantar o erro). Inerte fora do build Test.
Private Sub AreaTestHook(site As String, key As String)
    If BUILD_CONFIG = "Test" Then
        If Len(Util_FailThinSyncForTest) > 0 Then
            If Util_FailThinSyncForTest = "area-" & site & ":" & key Then
                Util_FailThinSyncForTest = ""
                Err.Raise vbObjectError + 9993, "ThinQQ", "Falha simulada (teste) " & site & " " & key
            End If
        End If
    End If
End Sub

Private Function EndListHas(list As String, key As String) As Boolean
    EndListHas = (InStr("," & list & ",", "," & key & ",") > 0)
End Function

Private Function EndListAdd(list As String, key As String) As String
    If EndListHas(list, key) Then EndListAdd = list: Exit Function
    If Len(list) = 0 Then EndListAdd = key Else EndListAdd = list & "," & key
End Function

Private Function EndListRemove(list As String, key As String) As String
    Dim s As String
    s = Replace("," & list & ",", "," & key & ",", ",")
    If Len(s) <= 1 Then Exit Function
    EndListRemove = Mid$(s, 2, Len(s) - 2)
End Function

' Rotulos das PONTAS (1o e ultimo ponto de cada tqq_lbl_j, empurrados para dentro da faixa: Right no primeiro, Left no ultimo) que encostariam
' no rotulo vizinho da mesma faixa (2o / penultimo, centrado) ficam ocultos, como os segmentos finos da cascata (na demonstracao, 12 meses de
' 5 digitos davam "3.5713.418"). Sobreposicao pela escala atual: na horizontal, deslocamento do rotulo da ponta (AREA_SIDE_OFFSET) + largura
' estimada dele + meia largura do vizinho + 1 pt > largura da fatia (com 2 categorias, os dois sao de ponta); na vertical, os centros (meio
' da faixa, Mid_j, na escala do eixo) a menos de AREA_TEXT_HEIGHT. Larguras pelo texto do formato atual (Labels_EstimateWidth do valor da
' celula ligada: Rnd_j com a soma consistente, senao Val_j). Valor zero (rotulo vazio) fica como esta.
' Disciplina da Fase 2: o ThinQQ so religa o que ELE ocultou (pares j.i na opcao "thinlbl", a mesma da cascata; "Alterar dados" e religar os
' rotulos a limpam): um rotulo que falta e nao esta na lista foi apagado pelo usuario e fica apagado. Transacional por ponto: as variaveis do
' ponto sao zeradas antes de cada leitura, um par entra na lista so se ocultar deu certo e sai so se religar deu certo (uma falha no meio de
' religar oculta de novo); a lista e gravada tambem na saida por erro. So escreve o que mudou. Sai cedo durante Labels_LinkPoints. Nunca
' propaga erro.
Public Sub Area_SyncEndLabels(cht As Chart)
    Dim blk As Range, body As Range, n As Long, ns As Long, j As Long, s As Series, fmt As String, slot As Double, mn As Double, mx As Double
    Dim list0 As String, list As String, e As Long, i As Long, nb As Long, key As String, txt As String, txtN As String, w As Double, wn As Double
    Dim y As Double, yn As Double, hide As Boolean, has As Boolean, readOk As Boolean, vc As Long, c As Range, cn As Range, sheetRef As String
    Dim fc As Long, pos As Long, need As Double, en As Long, ed As String, rs As Boolean
    On Error GoTo Done
    If Labels_IsLinking() Then Exit Sub
    If Kinds_Family(Registry_ChartType(cht)) <> "area" Then Exit Sub
    list0 = Registry_GetOption(cht, AREA_THIN_OPT, ""): list = list0
    If Registry_GetOption(cht, "labels", "1") <> "1" Then GoTo Done
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then GoTo Done
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    If n < 2 Then GoTo Done
    ns = Area_SeriesCount(blk)
    slot = cht.PlotArea.InsideWidth / n
    mn = cht.Axes(xlValue, xlPrimary).MinimumScale: mx = cht.Axes(xlValue, xlPrimary).MaximumScale
    If mx <= mn Or slot <= 0 Then GoTo Done
    fmt = Labels_SegmentFormat(cht)
    sheetRef = "='" & Replace(body.Worksheet.Name, "'", "''") & "'!"
    rs = (Registry_GetOption(cht, "roundsum", "0") = "1")
    For j = 1 To ns
        Set s = Chart_FindSeries(cht, "tqq_lbl_" & j)
        If Not s Is Nothing Then
            vc = 0
            If rs Then vc = Round_RCol(blk, j)
            If vc = 0 Then vc = Area_ValCol(blk, j)
            fc = Style_ContrastFont(Style_DataSeriesColor(cht, j))
            For e = 1 To 2
                If e = 1 Then
                    i = 1: nb = 2: pos = xlLabelPositionRight
                Else
                    i = n: nb = n - 1: pos = xlLabelPositionLeft
                End If
                If i <= s.Points.Count Then
                    key = j & "." & i
                    hide = False: has = False: readOk = False: txt = "": txtN = "": w = 0: wn = 0: need = 0
                    Set c = body.Cells(i + 1, vc): Set cn = body.Cells(nb + 1, vc)
                    On Error Resume Next
                    AreaTestHook "read", key
                    If Err.Number = 0 Then
                        txt = Format$(Util_ToDouble(c.Value), fmt): txtN = Format$(Util_ToDouble(cn.Value), fmt)
                        y = Util_ToDouble(body.Cells(i + 1, Area_MidCol(blk, j)).Value)
                        yn = Util_ToDouble(body.Cells(nb + 1, Area_MidCol(blk, j)).Value)
                        has = s.Points(i).HasDataLabel
                    End If
                    readOk = (Err.Number = 0)
                    If readOk And Len(txt) > 0 Then
                        w = Labels_EstimateWidth(txt)
                        If Len(txtN) > 0 Then wn = Labels_EstimateWidth(txtN)
                        If n = 2 Then need = 2 * AREA_SIDE_OFFSET + w + wn + 1 Else need = AREA_SIDE_OFFSET + w + wn / 2 + 1
                        hide = (need > slot And wn > 0 And Abs(y - yn) * cht.PlotArea.InsideHeight / (mx - mn) < AREA_TEXT_HEIGHT)
                        If hide Then
                            If has Then
                                s.Points(i).HasDataLabel = False
                                If Err.Number = 0 Then list = EndListAdd(list, key)
                            End If
                        ElseIf Not has And EndListHas(list, key) Then
                            AreaTestHook "link", key
                            If Err.Number = 0 Then Labels_LinkOnePoint s, i, sheetRef & c.Address(True, True), c.Value, fmt, pos, fc
                            If Err.Number = 0 Then
                                list = EndListRemove(list, key)
                            Else
                                en = Err.Number: ed = Err.Description: Err.Clear
                                s.Points(i).HasDataLabel = False                    ' religar falhou no meio: oculto de novo, continua na lista
                                Err.Raise en, "ThinQQ", ed
                            End If
                        End If
                    End If
                    If Err.Number <> 0 Then Errors_Log "Area_SyncEndLabels" & vbTab & key & vbTab & Err.Number & vbTab & Err.Description
                    Err.Clear
                    On Error GoTo Done
                End If
            Next
        End If
    Next
Done:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1                                 ' encerra um erro ativo: sem isso o Resume Next abaixo nao protegeria a gravacao da lista
    On Error Resume Next
    If list <> list0 Then Registry_SetOption cht, AREA_THIN_OPT, list   ' tambem na saida por erro: o que ja foi ocultado fica registrado
    If en <> 0 Then Errors_Log "Area_SyncEndLabels" & vbTab & en & vbTab & ed
End Sub

' Totais: a auxiliar tqq_total em TotalPos com o rotulo ligado a Total, em negrito, acima/abaixo pelo sinal. As colunas 1..3
' do bloco da area tem o mesmo significado das do bloco empilhado, entao o trabalho e o de Stacked_SetTotals (que decide por
' Kinds_Supports(tipo, "totals")).
Public Sub Area_SetTotals(cht As Chart, turnOn As Boolean)
    Stacked_SetTotals cht, turnOn
End Sub
