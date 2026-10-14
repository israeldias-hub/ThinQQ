Attribute VB_Name = "modCoreApi"
Option Explicit
' API pública chamada pelo loader via Application.Run. Nunca propaga erro.
Private mLoaderPath As String
Private mApplyBusy As Boolean      ' guarda de reentrada de Core_ApplyPendingLoader
Private mLastNotice As String      ' última mensagem de Core_Notify (usada pelos testes)
Private mNoticeCount As Long        ' chamadas a Core_Notify desde Core_ClearLastNotice (testes)
Private mMismatchSeen As Collection ' ids de gráficos já avisados de tabela de origem com outro tamanho
Private mClipSeen As Collection     ' ids de gráficos já avisados de dados acima da escala fixa (Mesma escala)
Private mCopySeen As Collection     ' ids de gráficos já avisados de cópia em outra planilha (planilha copiada)
Private mTestFamily As String       ' gancho de teste: familia de despacho forcada (so no build Test); "" = a familia real do tipo
Private mLastConfirm As String      ' ultima pergunta de Core_Confirm (usada pelos testes)
Private mTestRoundDec As Variant    ' gancho de teste: respostas do Arredondar intervalo no build Test (Core_TestSetRoundRangeAnswers)
Private mTestRoundDest As Range
Private mTestBreakAnswer As Variant  ' gancho de teste: resposta do InputBox do Corte de eixo no build Test (Core_TestSetAxisBreakAnswer)
' Familias de tipo que tem despacho de paleta em ApplyOption (a lista usada ANTES de gravar a opcao; o Select Case do despacho a repete).
Private Const PALETTE_FAMILIES As String = ",stacked,combo,waterfall,area,pie,mekko,gantt,butterfly,scatter,waterfallpair,"

' Última mensagem enviada ao usuário por Core_Notify ("" se nenhuma). Gancho de teste (inofensivo em Release).
Public Function Core_LastNotice() As String
    Core_LastNotice = mLastNotice
End Function

Public Function Core_NoticeCount() As Long
    Core_NoticeCount = mNoticeCount
End Function

Public Sub Core_ClearLastNotice()
    mLastNotice = ""
    mNoticeCount = 0
    mLastConfirm = ""
End Sub

Public Function Core_Ping() As String
    Core_Ping = "pong"
End Function

Public Function Core_Version() As String
    Core_Version = APP_VERSION
End Function

Public Function Core_LoaderPath() As String
    Core_LoaderPath = mLoaderPath
End Function

Public Sub Core_OnLoaded()
    On Error Resume Next
    Watchers_Start
End Sub

Public Sub Core_OnUnload()
    On Error Resume Next
    Watchers_Stop
    Set mMismatchSeen = Nothing
    Set mClipSeen = Nothing
    Set mCopySeen = Nothing
End Sub

Public Sub Core_RegisterLoader(loaderName As String, loaderFullPath As String)
    On Error Resume Next
    Watchers_SetLoader loaderName
    mLoaderPath = loaderFullPath
End Sub

' Mensagem ao usuário (suprimida em build de teste para não travar o Excel invisível).
' logIt=False quando o chamador já registrou a mensagem (Errors_Handle).
Public Sub Core_Notify(msg As String, Optional style As VbMsgBoxStyle = vbInformation, Optional logIt As Boolean = True)
    On Error Resume Next
    mLastNotice = msg
    mNoticeCount = mNoticeCount + 1
    If logIt Then Errors_Log "notify" & vbTab & Replace(msg, vbLf, " | ")
    If BUILD_CONFIG = "Test" Then Exit Sub
    MsgBox msg, style, APP_NAME
End Sub

' Pergunta Sim/Não ao usuário; True só com "Sim". Como Core_Notify, é o único ponto de MsgBox de pergunta do core: no build de teste só
' registra e responde "Não" (um modal travaria o Excel invisível). Nunca propaga erro (falha = "Não").
Public Function Core_Confirm(msg As String) As Boolean
    On Error Resume Next
    mLastConfirm = msg
    Errors_Log "confirm" & vbTab & Replace(msg, vbLf, " | ")
    If BUILD_CONFIG = "Test" Then Exit Function
    Core_Confirm = (MsgBox(msg, vbQuestion + vbYesNo, APP_NAME) = vbYes)
End Function

' Última pergunta feita por Core_Confirm ("" se nenhuma desde Core_ClearLastNotice). Gancho de teste (inofensivo em Release).
Public Function Core_LastConfirm() As String
    On Error Resume Next
    Core_LastConfirm = mLastConfirm
End Function

Public Function Core_ActiveChart() As Chart
    Dim c As Chart
    On Error Resume Next
    Set c = Application.ActiveChart
    If c Is Nothing Then Exit Function
    If Registry_IsThinQQChart(c) Then Set Core_ActiveChart = c
End Function

Public Function Core_HasChart() As Boolean
    On Error Resume Next
    Core_HasChart = Not Core_ActiveChart() Is Nothing
End Function

' True se o gráfico ThinQQ ativo suporta o recurso (matriz de capacidades Kinds_Supports). "changesource" vale para
' qualquer gráfico ThinQQ. False sem gráfico ThinQQ ativo e em qualquer dúvida; nunca propaga erro.
Public Function Core_Supports(feature As String) As Boolean
    Dim c As Chart
    On Error GoTo Fail
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Function
    If feature = "changesource" Then
        Core_Supports = True                                  ' reservado: nenhum controle da faixa consulta "changesource" hoje (Alterar dados vale para todo grafico ThinQQ)
    ElseIf feature = "breakoff" Then
        Core_Supports = Kinds_SupportsChart(c, "break") And Len(Registry_GetOption(c, "vbreak", "")) > 0      ' Remover corte de eixo: so com corte
    Else
        Core_Supports = Kinds_SupportsChart(c, feature)       ' inclui a regra da legenda da cascata (so com varias series)
    End If
    Exit Function
Fail:
    Core_Supports = False
End Function

Public Sub Core_Insert(kind As String)
    On Error GoTo Fail
    Dim data As clsChartData, msg As String, co As ChartObject, ws As Worksheet, prevSU As Boolean
    Dim en As Long, ed As String
    prevSU = Application.ScreenUpdating
    If Not Kinds_IsValid(kind) Then
        Core_Notify "Tipo de gráfico desconhecido: " & kind, vbExclamation
        Exit Sub
    End If
    If TypeName(Selection) <> "Range" Then
        Core_Notify "Selecione a tabela de dados (cabeçalho + categorias) antes de inserir o gráfico."
        Exit Sub
    End If
    Set data = ChartData_FromRange(Selection)
    msg = data.Validate()
    If Len(msg) > 0 Then Core_Notify msg, vbExclamation: Exit Sub
    Set ws = data.Source.Worksheet
    If ws.ProtectContents Then
        Core_Notify "A planilha está protegida. Desproteja a planilha para inserir o gráfico.", vbExclamation
        Exit Sub
    End If
    Application.ScreenUpdating = False
    CalcSheet_Cleanup ws.Parent
    Set co = Kinds_Insert(ws, data, kind)
    ApplySourceNumFmt co.Chart, kind, data           ' formato inicial dos rotulos pelo formato da tabela (1a celula de valor)
    Application.ScreenUpdating = prevSU
    Watchers_Rehook
    co.Select
    Core_SafeRefresh co.Chart
    NotifyPieIgnoredColumns kind, data
    NotifyMekkoNegatives kind, data
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    Application.ScreenUpdating = prevSU
    If en = vbObjectError + 7200 Then
        ' Tipo válido ainda sem implementação nesta versão: aviso ao usuário, não é falha (nada foi criado).
        Core_Notify ed, vbInformation
    ElseIf en >= vbObjectError + 6007 And en <= vbObjectError + 6099 Then
        ' Erro de VALIDACAO DOS DADOS de um tipo novo (6007 cascata, 6008 combinacao, 6009 Mekko, 6010 Gantt): aviso ao
        ' usuario, nao e falha (nada foi criado). Os 6001..6006 sao falhas estruturais, tratadas por Errors_Handle.
        Core_Notify ed, vbExclamation
    Else
        Errors_HandleErr "Core_Insert(" & kind & ")", en, ed
    End If
End Sub

' Chave de numfmt que corresponde a um codigo de formato de celula (ingles): percentual -> pct0 (sem casas) ou pct1; escala de milhares
' ("#,##0,") -> mil e de milhoes ("0.0,,") -> mi; 1 casa -> dec1; 2 ou mais -> dec2; senao (Geral, inteiro, data, texto) int. So a primeira
' secao conta, limpa por Style_CleanFormatSection (textos, colchetes como [Red] e [$R$-416], literais).
Private Function NumFmtKeyFromCode(code As String) As String
    Dim clean As String, i As Long, ch As String, p As Long, d As Long, last As Long, commas As Long
    On Error GoTo Fail
    NumFmtKeyFromCode = "int"
    If Len(code) = 0 Or code = "General" Or code = "@" Then Exit Function
    clean = Style_CleanFormatSection(code)
    If clean Like "*[dDmMyYhHsS]*" Then Exit Function              ' data ou hora
    p = InStr(clean, ".")
    If p > 0 Then
        For i = p + 1 To Len(clean)
            ch = Mid$(clean, i, 1)
            If ch = "0" Or ch = "#" Or ch = "?" Then d = d + 1 Else Exit For
        Next
    End If
    ' escala: virgulas logo depois do ULTIMO digito do formato ("#,##0," = milhares; "0.0,," = milhoes)
    For i = 1 To Len(clean)
        ch = Mid$(clean, i, 1)
        If ch = "0" Or ch = "#" Or ch = "?" Then last = i
    Next
    If last > 0 Then
        For i = last + 1 To Len(clean)
            If Mid$(clean, i, 1) = "," Then commas = commas + 1 Else Exit For
        Next
    End If
    If InStr(clean, "%") > 0 Then
        NumFmtKeyFromCode = IIf(d = 0, "pct0", "pct1")
    ElseIf commas >= 2 Then
        NumFmtKeyFromCode = "mi"
    ElseIf commas = 1 Then
        NumFmtKeyFromCode = "mil"
    ElseIf d = 1 Then
        NumFmtKeyFromCode = "dec1"
    ElseIf d >= 2 Then
        NumFmtKeyFromCode = "dec2"
    End If
    Exit Function
Fail:
    NumFmtKeyFromCode = "int"
End Function

' Formato inicial (numfmt) do grafico recem-inserido pelo formato da primeira celula NUMERICA da primeira coluna de valores (texto, vazio e erro
' sao pulados; a primeira celula se nenhuma for numerica): uma tabela com 1 casa decimal, em milhares ou em % nao vira rotulos inteiros. So nos
' tipos com numfmt na matriz e so quando o tipo nasceu com "int" (o Mekko (%) escolhe pct0 para as participacoes). Pizza e rosca: pct0/pct1
' MOSTRAM A PARTICIPACAO de cada fatia no total; so ficam quando a origem em % ja e participacao (a coluna soma ~100%); senao o valor em
' decimais (dec2). O usuario continua trocando o formato depois. Uma falha aqui nao desfaz a insercao: fica registrada e o grafico mantem
' "int". Nunca propaga erro.
Private Sub ApplySourceNumFmt(cht As Chart, kind As String, data As clsChartData)
    Dim i As Long, c As Range, v As Variant, key As String, sum As Double
    On Error GoTo Fail
    If Not Kinds_Supports(kind, "numfmt") Then Exit Sub
    If Registry_GetOption(cht, "numfmt", "int") <> "int" Then Exit Sub
    For i = 1 To data.CategoryCount
        v = data.ValueCell(i, 1).Value2
        If Not IsError(v) And Not IsEmpty(v) And VarType(v) <> vbString And VarType(v) <> vbBoolean Then
            If c Is Nothing Then Set c = data.ValueCell(i, 1)
            sum = sum + CDbl(v)
        End If
    Next
    If c Is Nothing Then Set c = data.ValueCell(1, 1)
    If IsNull(c.NumberFormat) Then Exit Sub
    key = NumFmtKeyFromCode(CStr(c.NumberFormat))
    If Kinds_Family(kind) = "pie" And (key = "pct0" Or key = "pct1") Then
        If Abs(sum - 1) > 0.01 Then key = "dec2"                    ' % que nao e participacao: o valor, nao a fatia do total
    End If
    If key <> "int" Then ApplyOption cht, "numfmt", key
    Exit Sub
Fail:
    Errors_Log "Core_Insert/numfmt inicial" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Pizza e rosca mostram uma serie so: a primeira coluna de valores. Com mais colunas na tabela as demais sao ignoradas e o usuario e avisado
' (informativo: o grafico ja foi criado). Nunca propaga erro.
Private Sub NotifyPieIgnoredColumns(kind As String, data As clsChartData)
    On Error Resume Next
    If Kinds_Family(kind) <> "pie" Then Exit Sub
    If data.SeriesCount < 2 Then Exit Sub
    Core_Notify "A " & IIf(kind = "doughnut", "rosca", "pizza") & " mostra só a primeira coluna de valores (" & data.SeriesName(1) & "). As outras " & (data.SeriesCount - 1) & " coluna(s) da tabela foram ignoradas.", vbInformation
End Sub

' Mekko (%): as areas so mostram participacoes positivas; valores negativos da tabela entram como zero (Mekko_BuildBlock). Avisa depois de
' criar (informativo, como NotifyPieIgnoredColumns; no build de teste so registra). Nunca propaga erro.
Private Sub NotifyMekkoNegatives(kind As String, data As clsChartData)
    Dim j As Long, n As Long
    On Error Resume Next
    If kind <> "mekko" Then Exit Sub
    For j = 1 To data.SeriesCount
        n = n + Application.WorksheetFunction.CountIf(data.SeriesValueCells(j), "<0")
    Next
    If n > 0 Then Core_Notify "Mekko (%): os valores negativos da tabela entram como zero (as áreas mostram só as partes positivas).", vbInformation
End Sub

Public Function Core_GetToggle(key As String) As Boolean
    Dim c As Chart
    On Error Resume Next
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Function
    Core_GetToggle = (Registry_GetOption(c, key, "0") = "1")
End Function

' Aplica o toggle no gráfico informado (sem depender da seleção). Propaga erros. Chaves fora da lista de toggles saem sem fazer nada
' (silenciosas, como sempre); para uma chave da lista a MATRIZ de capacidades decide: tipo sem o recurso => erro 7002 (Core_SetToggle mostra o
' aviso e nada é gravado); com o recurso, despacha pela família do tipo.
Private Sub ApplyToggle(c As Chart, key As String, value As Boolean)
    Dim kind As String, fam As String
    Select Case key
        Case "totals", "labels", "connectors", "ind100", "vline", "secondary", "pieout", "trend", "quad", "roundsum", "pairdiff", "pairabs"
        Case Else: Exit Sub
    End Select
    kind = Registry_ChartType(c)
    If Not Kinds_Supports(kind, key) Then Err.Raise vbObjectError + 7002, "Core", Kinds_NotApplicableText(key)
    fam = DispatchFamily(c)
    ' Cada familia tem o seu Case: uma familia sem Case (a matriz diz que o tipo suporta, o despacho nao sabe fazer) e recusada com o aviso 7002,
    ' nunca despachada para o padrao da Fase 1 (colunas), que mexeria no grafico errado.
    Select Case key
        Case "totals"
            Select Case fam
                Case "area": Area_SetTotals c, value
                Case "mekko": Mekko_SetTotals c, value
                Case "butterfly": Bfly_SetTotals c, value             ' borboleta: auxiliar XY tqq_total no fim da barra direita
                Case "waterfallpair": Wfp_SetTotalPct c, value          ' cascata comparativa: o percentual do 1o total nas barras de total
                Case "stacked", "combo": Stacked_SetTotals c, value     ' colunas, barras (empilhadas e 100%) e a combinacao (so os totais das colunas)
                Case Else: RefuseUndispatched key, c, fam
            End Select
        Case "labels"
            Select Case fam
                Case "waterfall": Waterfall_SetLabels c, value
                Case "area": Area_SetLabels c, value
                Case "mekko": Mekko_SetLabels c, value
                Case "gantt": Gantt_SetLabels c, value
                Case "pie": Pie_SetLabels c, value
                Case "butterfly": Bfly_SetLabels c, value             ' borboleta: rotulos nativos dos dois lados (esquerdo absoluto)
                Case "scatter": Scatter_SetLabels c, value            ' dispersao e bolhas: o nome de cada ponto (ligado a coluna Nome do bloco)
                Case "waterfallpair": Wfp_SetLabels c, value          ' cascata comparativa: degraus no centro, barras grandes acima
                Case "stacked", "combo": Stacked_SetLabels c, value     ' colunas, barras, linha e a combinacao
                Case Else: RefuseUndispatched key, c, fam
            End Select
        Case "connectors"
            Select Case fam
                Case "waterfall": Ann_SetWaterfallConnectors c, value
                Case "waterfallpair": Wfp_SetConnectors c, value     ' cascata comparativa: auxiliares XY com barra de erro, por serie
                Case "stacked", "combo": Stacked_SetConnectors c, value
                Case Else: RefuseUndispatched key, c, fam
            End Select
        Case "ind100"
            Deco_SetInd100 c, value
        Case "vline"
            Deco_SetValueLine c, value                      ' recusa (7002) em grafico antigo, sem os dados da linha de valor
        Case "secondary"
            Combo_SetSecondary c, value
        Case "pieout"
            Pie_SetOutside c, value
        Case "trend"
            If fam = "scatter" Then Scatter_SetTrend c, value Else RefuseUndispatched key, c, fam     ' linha de tendencia linear por grupo
        Case "quad"
            If fam = "scatter" Then Scatter_SetQuad c, value Else RefuseUndispatched key, c, fam      ' quadrantes: eixos cruzando nas medias
        Case "pairdiff"
            If fam = "waterfallpair" Then Wfp_SetArrows c, value Else RefuseUndispatched key, c, fam      ' setas de variacao entre as barras do par
        Case "pairabs"
            If fam = "waterfallpair" Then Wfp_SetArrowMode c, value Else RefuseUndispatched key, c, fam   ' variacao em valor em vez de percentual
        Case "roundsum"
            Select Case fam
                Case "stacked", "area": Round_SetLabels c, value     ' rotulos com soma consistente (maior resto nas colunas Rnd_j do bloco)
                Case Else: RefuseUndispatched key, c, fam
            End Select
    End Select
End Sub

' Familia de despacho do grafico (a do tipo; o gancho de teste Core_TestSetFamilyOverride a substitui).
Private Function DispatchFamily(c As Chart) As String
    If Len(mTestFamily) > 0 Then DispatchFamily = mTestFamily Else DispatchFamily = Kinds_Family(Registry_ChartType(c))
End Function

' O despacho nao tem a familia do tipo para este recurso (a matriz de capacidades e o despacho divergiram: tipo novo na matriz sem Case aqui):
' registra a divergencia no log e recusa com 7002 (o mesmo aviso de "nao se aplica"), em vez de cair no padrao de outra familia.
Private Sub RefuseUndispatched(key As String, c As Chart, fam As String)
    Errors_Log "despacho sem familia" & vbTab & key & vbTab & Registry_ChartType(c) & vbTab & fam
    Err.Raise vbObjectError + 7002, "Core", Kinds_NotApplicableText(key)
End Sub

' Valida o valor de uma opção. Retorna "" se válido, senão a mensagem de erro.
Private Function ValidateOption(key As String, value As String) As String
    Select Case key
        Case "numfmt"
            Select Case value
                Case "int", "dec1", "dec2", "pct0", "pct1", "mil", "mi"
                Case Else: ValidateOption = "Formato numérico desconhecido: " & value
            End Select
        Case "palette"
            Select Case value
                Case "padrao", "cinza", "tema"
                Case Else: ValidateOption = "Paleta desconhecida: " & value
            End Select
        Case "legend"
            Select Case value
                Case "right", "none", "lineend", "inbar"
                Case Else: ValidateOption = "Opção de legenda desconhecida: " & value
            End Select
        Case "gap"
            If Not IsNumeric(value) Then
                ValidateOption = "O espaçamento deve ser um número entre 0 e 500."
            ElseIf Val(value) < 0 Or Val(value) > 500 Then
                ValidateOption = "O espaçamento deve ser um número entre 0 e 500."
            End If
        Case Else
            ValidateOption = "Opção desconhecida: " & key
    End Select
End Function

' Aplica a opção no gráfico informado (sem depender da seleção). Valida o valor e consulta a matriz de capacidades ANTES de gravar: opção que
' o tipo não tem (gap, legenda da cascata de uma série e do Gantt, formato do Gantt) => erro 7002 e nada é gravado. Propaga erros.
Private Sub ApplyOption(c As Chart, key As String, value As String)
    Dim msg As String, fam As String
    msg = ValidateOption(key, value)
    If Len(msg) > 0 Then Err.Raise vbObjectError + 7100, "Core", msg
    If Not Kinds_SupportsChart(c, key) Then Err.Raise vbObjectError + 7002, "Core", Kinds_NotApplicableText(key)
    If key = "legend" Then
        ' A legenda valida, aplica e SO ENTAO grava (Deco_SetLegend): uma recusa (7002) nao pode deixar legend=lineend gravado.
        ' As demais opcoes gravam antes de despachar porque o despacho le a opcao do registro (ex.: numfmt, em Labels_SegmentFormat).
        Deco_SetLegend c, value
        Exit Sub
    End If
    fam = DispatchFamily(c)
    ' paleta: a familia sem despacho e recusada ANTES de gravar (a recusa 7002 nunca deixa a opcao gravada)
    If key = "palette" Then
        If InStr(PALETTE_FAMILIES, "," & fam & ",") = 0 Then RefuseUndispatched key, c, fam
    End If
    Registry_SetOption c, key, value
    Select Case key
        Case "numfmt"
            Style_ApplyNumberFormat c, Style_NumberFormat(value)       ' despacha por familia (pizza: Pie_ApplyNumberFormat; Mekko: Mekko_ApplyNumberFormat)
            Chart_EnsureRoomAfterChange c               ' rotulos de total mais largos: refaz a folga dos totais e das setas
        Case "palette"
            Select Case fam
                Case "waterfall": Waterfall_ApplyColors c, value
                Case "area": Area_ApplyColors c, value
                Case "pie": Pie_ApplyColors c, value           ' cores por fatia (Style_ApplyPointColors) e fonte dos rotulos por contraste
                Case "mekko": Mekko_ApplyColors c, value        ' separadores de 1,5 pt e contraste dos rotulos (por ponto)
                Case "gantt": Gantt_ApplyColors c, value        ' barras pal(1) e marcos pal(0) (as series do Gantt sao todas tqq_*)
                Case "butterfly": Bfly_ApplyColors c, value     ' as duas series de dados; vaos invisiveis; contraste dos rotulos
                Case "scatter": Scatter_ApplyColors c, value    ' marcadores/bolhas por grupo, tendencias e contraste dos rotulos das bolhas
                Case "waterfallpair": Wfp_ApplyColors c, value              ' cascata comparativa: serie 1 pal(0), serie 2 pal(2), contraste dos rotulos
                Case "stacked", "combo": Style_ApplySeriesColors c, value
                Case Else: RefuseUndispatched key, c, fam       ' (inalcancavel: PALETTE_FAMILIES ja recusou antes de gravar)
            End Select
        Case "gap"
            ApplyGapWidth c, CLng(Val(value))
            If fam = "waterfallpair" Then Wfp_SyncGap c         ' cascata comparativa: meia-largura das barras (conectores) acompanha o espaco
    End Select
End Sub

' Espaco entre colunas/barras em TODOS os grupos de colunas ou barras do grafico (uma serie de colunas da combinacao levada ao eixo secundario
' forma um segundo grupo de colunas); os grupos de linha (combinacao, auxiliares) e XY (barras, Gantt) nao tem espaco e ficam como estao. Sem
' nenhum grupo de colunas/barras reconhecido, o primeiro grupo (como antes).
Private Sub ApplyGapWidth(c As Chart, gap As Long)
    Dim g As ChartGroup, n As Long
    For Each g In c.ChartGroups
        If IsColumnOrBarGroup(g) Then g.GapWidth = gap: n = n + 1
    Next
    If n = 0 Then c.ChartGroups(1).GapWidth = gap
End Sub

Private Function IsColumnOrBarGroup(g As ChartGroup) As Boolean
    If g.SeriesCollection.Count = 0 Then Exit Function
    Select Case g.SeriesCollection(1).ChartType
        Case xlColumnClustered, xlColumnStacked, xlColumnStacked100, xlBarClustered, xlBarStacked, xlBarStacked100
            IsColumnOrBarGroup = True
    End Select
End Function

Public Sub Core_SetToggle(key As String, value As Boolean)
    On Error GoTo Fail
    Dim c As Chart, msg As String
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    ApplyToggle c, key, value
    Core_SafeRefresh c
    If key = "vline" And value Then
        msg = Break_VLineHiddenText(c)                   ' aviso nao bloqueante: a media acima do corte de eixo nao aparece
        If Len(msg) > 0 Then Core_Notify msg, vbInformation
    End If
    Exit Sub
Fail:
    If Err.Number = vbObjectError + 7002 Then
        Core_Notify Err.Description, vbInformation       ' recusa de decoracao/opcao (erro 7002): aviso, nao e falha; nada e gravado
    Else
        Errors_Handle "Core_SetToggle(" & key & ")"
    End If
End Sub

Public Function Core_GetOption(key As String) As String
    Dim c As Chart
    On Error Resume Next
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Function
    Core_GetOption = Registry_GetOption(c, key, "")
End Function

Public Sub Core_SetOption(key As String, value As String)
    On Error GoTo Fail
    Dim c As Chart, msg As String
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    msg = ValidateOption(key, value)
    If Len(msg) > 0 Then Core_Notify msg, vbExclamation: Exit Sub
    ApplyOption c, key, value                            ' recusa pela matriz (7002) antes de gravar: o Fail: abaixo mostra o aviso
    Core_SafeRefresh c
    Exit Sub
Fail:
    If Err.Number = vbObjectError + 7002 Then
        Core_Notify Err.Description, vbInformation       ' recusa (ex.: "fim da linha" fora de linha/combinacao): aviso, nao e falha
    Else
        Errors_Handle "Core_SetOption(" & key & ")"
    End If
End Sub

' Setas (cagr, diff) e colchete (bracket, só no Gantt). A MATRIZ de tipos vem ANTES de qualquer outra verificação (inclusive a de poucas
' categorias e o formulário): tipo sem a anotação => erro 7001 com a mensagem do tipo (Ann_RefusalText), mostrada como aviso sem abrir formulário
' (uma pizza ou um Gantt de uma linha recebem a mensagem do tipo, não "pelo menos 2 categorias"). Depois valem as regras de índices: o colchete
' vale sobre UMA atividade (de = até): exige n >= 1 e de <= até; as setas exigem n >= 2 e de < até. Índices fora da regra abrem o formulário.
' Tipo de anotação desconhecido sai sem aviso ao usuário, mas fica registrado no log (as chaves desconhecidas de Core_SetToggle saem sem registro).
Public Sub Core_AddAnnotation(kind As String, Optional fromIdx As Long = 0, Optional toIdx As Long = 0)
    On Error GoTo Fail
    Dim c As Chart, n As Long, refusal As String, isBracket As Boolean
    Select Case kind
        Case "cagr", "diff", "bracket"
        Case Else
            Errors_Log "Core_AddAnnotation: tipo de anotacao desconhecido (ignorado)" & vbTab & kind       ' sem aviso ao usuario; a faixa so envia cagr, diff e bracket
            Exit Sub
    End Select
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    refusal = Ann_RefusalText(c, kind)
    If Len(refusal) > 0 Then Err.Raise vbObjectError + 7001, "Core_AddAnnotation", refusal
    isBracket = (kind = "bracket")
    n = Geom_ChartCategoryCount(c)
    If isBracket Then
        If n < 1 Then Core_Notify "O gráfico precisa de pelo menos 1 atividade.": Exit Sub
    ElseIf n < 2 Then
        Core_Notify "O gráfico precisa de pelo menos 2 categorias.": Exit Sub
    End If
    If fromIdx < 1 Or toIdx > n Or IIf(isBracket, fromIdx > toIdx, fromIdx >= toIdx) Then
        If Not Core_AskCategories(c, n, fromIdx, toIdx, isBracket) Then Exit Sub
    End If
    Select Case kind
        Case "cagr": Ann_AddCagr c, fromIdx, toIdx
        Case "diff": Ann_AddDifference c, fromIdx, toIdx
        Case Else: Ann_AddBracket c, fromIdx, toIdx
    End Select
    Exit Sub
Fail:
    If Err.Number = vbObjectError + 7001 Then
        ' Seta/colchete em tipo que não o suporta (recusa da matriz de tipos): aviso ao usuário, não é falha.
        Core_Notify Err.Description, vbInformation
    Else
        Errors_Handle "Core_AddAnnotation(" & kind & ")"
    End If
End Sub

' Abre o formulário de categorias (allowSame: o colchete aceita a mesma atividade nas duas pontas). Nunca abre em build de teste (modal
' travaria o Excel invisível).
Private Function Core_AskCategories(c As Chart, n As Long, ByRef fromIdx As Long, ByRef toIdx As Long, Optional allowSame As Boolean = False) As Boolean
    Dim frm As frmAnnotation, cats() As String, body As Range, i As Long
    If BUILD_CONFIG = "Test" Then Exit Function
    Set body = Block_Body(Registry_Block(c))
    ReDim cats(1 To n)
    For i = 1 To n
        cats(i) = Util_CellLabel(body.Cells(i + 1, 1))
    Next
    Set frm = New frmAnnotation
    frm.Prepare cats, 1, n, allowSame
    frm.Show
    If frm.Result = "ok" Then
        fromIdx = frm.FromIdx: toIdx = frm.ToIdx
        Core_AskCategories = True
    End If
    Unload frm
End Function

Public Sub Core_RemoveAnnotations()
    On Error GoTo Fail
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    Ann_RemoveAll c, False
    Exit Sub
Fail:
    Errors_Handle "Core_RemoveAnnotations"
End Sub

' Graficos ThinQQ que o usuario tem em maos: os ThinQQ de uma selecao de VARIOS objetos (Ctrl+clique; Selection e "DrawingObjects" e
' Selection.ShapeRange lista as formas) ou, senao, o grafico ThinQQ ativo. Colecao vazia sem nenhum. Nunca propaga erro.
Public Function Core_SelectedCharts() As Collection
    On Error Resume Next
    Dim out As Collection, c As Chart, sr As Object, i As Long, sh As Object, co As ChartObject
    Set out = New Collection
    Set Core_SelectedCharts = out
    If TypeName(Selection) = "DrawingObjects" Then
        Set sr = Selection.ShapeRange
        If Not sr Is Nothing Then
            For i = 1 To sr.Count
                Set sh = sr.Item(i)
                If sh.HasChart Then
                    Set co = Nothing
                    Set co = sh.Parent.ChartObjects(sh.Name)
                    If Not co Is Nothing Then
                        If Registry_IsThinQQChart(co.Chart) Then out.Add co.Chart
                    End If
                End If
                Err.Clear
            Next
        End If
    End If
    Err.Clear
    If out.Count > 0 Then Exit Function
    Set c = Core_ActiveChart()
    If Not c Is Nothing Then out.Add c
End Function

' Todos os graficos ThinQQ da planilha ativa (colecao vazia sem planilha). Nunca propaga erro.
Private Function SheetThinQQCharts() As Collection
    Dim out As Collection, co As ChartObject, ws As Object
    Set out = New Collection
    Set SheetThinQQCharts = out
    On Error Resume Next
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Function
    If TypeName(ws) <> "Worksheet" Then Exit Function
    For Each co In ws.ChartObjects
        If Left$(co.Name, Len(ID_PREFIX)) = ID_PREFIX Then
            If Registry_IsThinQQChart(co.Chart) Then out.Add co.Chart
        End If
        Err.Clear
    Next
End Function

' Mesma escala: os graficos ThinQQ selecionados (dois ou mais) ou, com menos de dois selecionados, TODOS os da planilha ativa (o botao
' fica sempre habilitado e diz o que fez). Aplica a uniao das escalas (Deco_SameScale) e informa quantos receberam e quantos foram
' ignorados pelo tipo (100%, Mekko (%), pizza e Gantt nao tem escala livre). A escala fica fixa: quem mudar muito os dados usa de novo.
Public Sub Core_SameScale()
    On Error GoTo Fail
    Dim coll As Collection, n As Long, skipped As Long, brk As Long, msg As String
    Set coll = Core_SelectedCharts()
    If coll.Count < 2 Then Set coll = SheetThinQQCharts()
    If coll.Count < 2 Then
        Core_Notify "Mesma escala: selecione dois ou mais gráficos do ThinQQ (ou nenhum, para usar todos os gráficos da planilha).", vbInformation
        Exit Sub
    End If
    n = Deco_SameScale(coll, skipped, brk)
    If n < 2 Then
        msg = "Mesma escala: são necessários dois ou mais gráficos com escala livre (100%, Mekko (%), pizza e Gantt não entram)."
        If brk > 0 Then msg = msg & " Gráficos com corte de eixo também ficam de fora."
        Core_Notify msg, vbInformation
        Exit Sub
    End If
    msg = "Mesma escala aplicada a " & n & " gráfico(s)."
    If skipped > 0 Then msg = msg & " " & skipped & " ignorado(s): tipo sem escala livre (100%, Mekko (%), pizza, Gantt)."
    If brk > 0 Then msg = msg & " " & brk & " com corte de eixo fora do grupo (corte de eixo e Mesma escala são excludentes)."
    Core_Notify msg, vbInformation
    Exit Sub
Fail:
    Errors_Handle "Core_SameScale"
End Sub

' Escala automatica: devolve ao automatico o eixo de valores do grafico ThinQQ ativo ou dos selecionados (sem nenhum, de todos os da
' planilha ativa, como a Mesma escala) e limpa a escala gravada. So os graficos com a escala fixa do ThinQQ (Deco_ResetScale): um eixo ajustado
' a mao pelo usuario fica como esta. Informa quantos graficos foram restaurados.
Public Sub Core_ResetScale()
    On Error GoTo Fail
    Dim coll As Collection, c As Variant, cht As Chart, n As Long, nFree As Long
    Set coll = Core_SelectedCharts()
    If coll.Count = 0 Then Set coll = SheetThinQQCharts()
    If coll.Count = 0 Then
        Core_Notify "Escala automática: selecione um gráfico do ThinQQ (ou nenhum, para usar todos os gráficos da planilha).", vbInformation
        Exit Sub
    End If
    For Each c In coll
        Set cht = c
        If Kinds_Supports(Registry_ChartType(cht), "samescale") Then
            nFree = nFree + 1
            If Deco_ResetScale(cht) Then n = n + 1
        End If
    Next
    If nFree = 0 Then
        Core_Notify "Escala automática: o gráfico não tem escala livre (100%, Mekko (%), pizza, Gantt).", vbInformation
    ElseIf n = 0 Then
        Core_Notify "Escala automática: nenhum gráfico tinha a escala fixa de 'Mesma escala' (um eixo ajustado à mão fica como está).", vbInformation
    Else
        Core_Notify "Escala automática restaurada em " & n & " gráfico(s).", vbInformation
    End If
    Exit Sub
Fail:
    Errors_Handle "Core_ResetScale"
End Sub

' Arredondar intervalo: grava, a partir da celula escolhida, formulas de maior resto que referenciam so a faixa selecionada (uma coluna ou
' uma linha de numeros). Portatil: nada depende do ThinQQ. Em build Test sai antes dos InputBox (modais travariam o Excel invisivel).
' O destino e validado (Round_ValidateDest: mesma pasta, cabe na planilha, sem sobrepor a origem) e celulas ja preenchidas pedem confirmacao.
Public Sub Core_RoundRange()
    On Error GoTo Fail
    Dim rng As Range, dest As Range, dec As Variant, msg As String, n As Long, out As Range, addr As String, tDec As Variant, tDest As Range
    tDec = mTestRoundDec: Set tDest = mTestRoundDest             ' gancho de teste de uso unico: consumido na entrada (nao vaza de uma saida antecipada)
    Set mTestRoundDest = Nothing: mTestRoundDec = Empty
    If TypeName(Selection) <> "Range" Then
        Core_Notify "Selecione uma única coluna ou linha só com números (pelo menos 2).", vbExclamation
        Exit Sub
    End If
    Set rng = Selection
    msg = Round_ValidateRange(rng)
    If Len(msg) > 0 Then Core_Notify msg, vbExclamation: Exit Sub
    If BUILD_CONFIG = "Test" Then
        ' build Test: sem InputBox (modais travariam o Excel invisivel); as respostas vem do gancho Core_TestSetRoundRangeAnswers (uso unico)
        If tDest Is Nothing Then Exit Sub
        dec = tDec: Set dest = tDest
    Else
        dec = Application.InputBox("Casas decimais do resultado (0 = inteiros; 1 ou 2 = decimais; -3 = milhares):", APP_NAME, 0, Type:=1)
    End If
    If VarType(dec) = vbBoolean Then Exit Sub
    If Not IsNumeric(dec) Then dec = 0.5                       ' cai na mensagem abaixo
    If dec <> Int(dec) Or dec < -15 Or dec > 15 Then
        Core_Notify "Informe um número inteiro de casas decimais (de -15 a 15).", vbExclamation
        Exit Sub
    End If
    If dest Is Nothing Then
        On Error Resume Next
        Set dest = Application.InputBox("Célula onde o resultado começa (cada célula recebe uma fórmula que referencia só os valores de origem):", APP_NAME, Type:=8)
        On Error GoTo Fail
    End If
    If dest Is Nothing Then Exit Sub
    msg = Round_ValidateDest(rng, dest.Cells(1, 1))
    If Len(msg) > 0 Then Core_Notify msg, vbExclamation: Exit Sub
    n = rng.Cells.Count
    If rng.Columns.Count = 1 Then Set out = dest.Cells(1, 1).Resize(n, 1) Else Set out = dest.Cells(1, 1).Resize(1, n)
    addr = out.Address(False, False)
    If Not out.Worksheet Is rng.Worksheet Then addr = "'" & Replace(out.Worksheet.Name, "'", "''") & "'!" & addr
    If Application.WorksheetFunction.CountA(out) > 0 Then
        If Not Core_Confirm("As células " & addr & " já têm conteúdo. Substituir? Esta ação não pode ser desfeita.") Then Exit Sub
    End If
    Round_WriteFormulas rng, dest.Cells(1, 1), CLng(dec)
    Core_Notify "Fórmulas de arredondamento com soma consistente gravadas em " & addr & ".", vbInformation
    Exit Sub
Fail:
    If Err.Number = vbObjectError + 7002 Then
        Core_Notify Err.Description, vbExclamation       ' recusa de Round_WriteFormulas (tamanho ou destino): aviso, nada gravado
    Else
        Errors_Handle "Core_RoundRange"
    End If
End Sub

' Corte de eixo: pede o valor (sugestao: Break_Suggest) e aplica (Core_SetAxisBreakAt). Tipo sem corte: aviso antes de perguntar. Valor <= 0:
' o aviso do limite (remover e Core_RemoveAxisBreak). Em build Test sai antes do InputBox; a resposta pode vir do gancho
' Core_TestSetAxisBreakAnswer (uso unico, consumido logo na entrada: nao vaza para a chamada seguinte em nenhum caminho).
Public Sub Core_SetAxisBreak()
    On Error GoTo Fail
    Dim c As Chart, v As Variant, ans As Variant
    ans = mTestBreakAnswer: mTestBreakAnswer = Empty
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    If Not Kinds_SupportsChart(c, "break") Then
        Core_Notify Kinds_NotApplicableText("break"), vbInformation
        Exit Sub
    End If
    If BUILD_CONFIG = "Test" Then
        If IsEmpty(ans) Then Exit Sub                     ' build Test: sem InputBox (modal travaria o Excel invisivel)
        v = ans
    Else
        v = Application.InputBox("Corte de eixo: valor onde o eixo é cortado (as colunas acima dele ficam cortadas e o total continua no rótulo):", APP_NAME, CStr(Break_Suggest(c)), Type:=1)
    End If
    If VarType(v) = vbBoolean Then Exit Sub               ' Cancelar
    If Not IsNumeric(v) Then v = 0
    If CDbl(v) <= 0 Then
        Core_Notify BREAK_ERR_LIMIT, vbInformation
        Exit Sub
    End If
    Core_SetAxisBreakAt CDbl(v)
    Exit Sub
Fail:
    Errors_Handle "Core_SetAxisBreak"
End Sub

' Aplica o corte em limit no grafico ativo (limit <= 0 remove). Recusas (tipo sem corte, limite fora de 0 < limit < maior topo): aviso, nada
' gravado. Avisos nao bloqueantes (uma caixa so): o grafico saiu do grupo de Mesma escala (corte e Mesma escala sao excludentes) e a linha de
' valor ficou acima do corte (some).
Public Sub Core_SetAxisBreakAt(limit As Double)
    On Error GoTo Fail
    Dim c As Chart, hadScale As Boolean, msg As String
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    hadScale = (Len(Registry_GetOption(c, "vmax", "")) > 0)
    Break_Set c, limit
    Core_SafeRefresh c
    If hadScale And limit > 0 Then
        If Len(Registry_GetOption(c, "vmax", "")) = 0 Then msg = "O gráfico saiu do grupo de Mesma escala: corte de eixo e Mesma escala são excludentes."
    End If
    msg = Trim$(msg & " " & Break_VLineHiddenText(c))
    If Len(msg) > 0 Then Core_Notify msg, vbInformation
    Exit Sub
Fail:
    If Err.Number = vbObjectError + 7002 Or Err.Number = vbObjectError + 7003 Then
        Core_Notify Err.Description, vbInformation       ' recusa (tipo sem corte, limite invalido): aviso, nada gravado
    Else
        Errors_Handle "Core_SetAxisBreakAt"
    End If
End Sub

' Remove o corte de eixo do grafico ativo (sem corte: nada).
Public Sub Core_RemoveAxisBreak()
    On Error GoTo Fail
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    If Len(Registry_GetOption(c, "vbreak", "")) = 0 Then Exit Sub
    Break_Remove c
    Core_SafeRefresh c
    Exit Sub
Fail:
    Errors_Handle "Core_RemoveAxisBreak"
End Sub

' True quando o bloco do gráfico (uma linha de fórmula por categoria) não tem o mesmo número de categorias
' que o gráfico mostra agora: linhas inseridas/removidas na tabela de origem ajustam as referências diretas
' das séries (e as categorias), mas não o bloco. Nunca propaga erro (False em dúvida).
Public Function Core_ChartSourceMismatch(cht As Chart) As Boolean
    Dim blk As Range, v As Variant, k As Long, nOpt As String, expected As Long
    On Error GoTo Fail
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    If cht.SeriesCollection.Count = 0 Then Exit Function
    k = Util_SeriesCategoryRefCount(cht.SeriesCollection(1), Registry_Workbook(cht))
    If k < 0 Then
        v = cht.SeriesCollection(1).XValues
        If IsArray(v) Then k = UBound(v) - LBound(v) + 1 Else k = 1
    End If
    ' Com a opção "n" (número de categorias, gravada pelos Insert), o esperado vem do tipo (Mekko: 5 pontos por
    ' categoria); sem ela (gráfico de versão antiga), o bloco tem meta + cabeçalho + 1 linha por categoria.
    nOpt = Registry_GetOption(cht, "n", "")
    If Len(nOpt) > 0 Then
        expected = Kinds_ExpectedCategoryRefs(Registry_ChartType(cht), CLng(Val(nOpt)))
    Else
        expected = blk.Rows.Count - 2
    End If
    Core_ChartSourceMismatch = (k <> expected)
    ' Mekko: as series apontam para o BLOCO (sempre 5n pontos) e nao mudam com a tabela; o tamanho dela vem das referencias das
    ' categorias no bloco (linhas inseridas no meio da tabela, ou apagadas, mudam o numero de linhas entre a primeira e a ultima)
    If Not Core_ChartSourceMismatch And Len(nOpt) > 0 Then
        If Kinds_Family(Registry_ChartType(cht)) = "mekko" Then Core_ChartSourceMismatch = (Mekko_SourceRowsSpanned(cht) <> CLng(Val(nOpt)))
        ' dispersao/bolhas: o mesmo (as series apontam para o bloco; o tamanho da tabela so aparece nas referencias da coluna X do bloco)
        If Kinds_Family(Registry_ChartType(cht)) = "scatter" Then Core_ChartSourceMismatch = (Scatter_SourceRowsSpanned(cht) <> CLng(Val(nOpt)))
        If Kinds_Family(Registry_ChartType(cht)) = "waterfallpair" Then Core_ChartSourceMismatch = (Wfp_SourceRowsSpanned(cht) <> CLng(Val(nOpt)))
    End If
    Exit Function
Fail:
    Core_ChartSourceMismatch = False
End Function

' Avisa uma única vez por gráfico e por sessão (reiniciado em Core_OnUnload).
Private Sub NotifySourceMismatchOnce(cht As Chart)
    Dim id As String
    On Error Resume Next
    id = Registry_IdFromChart(cht)
    If Len(id) = 0 Then Exit Sub
    If Not Core_ChartSourceMismatch(cht) Then Exit Sub
    If mMismatchSeen Is Nothing Then Set mMismatchSeen = New Collection
    Err.Clear
    mMismatchSeen.Add id, id
    If Err.Number <> 0 Then Err.Clear: Exit Sub       ' chave repetida: já avisado
    Core_Notify "A tabela de origem mudou de tamanho. Use 'Alterar dados' para reconstruir o gráfico (ou reinsira o gráfico).", vbExclamation
End Sub

' Planilha copiada (Mover ou copiar > Criar uma cópia): o gráfico copiado mantém o nome tqq_<id> e passa a dividir id, bloco e opções com o
' original, com as séries auxiliares ainda na tabela original. Chamado ao observar a planilha (Watchers_Rehook): se outro ChartObject da pasta
' tem o mesmo nome, avisa uma única vez por id e por sessão para usar Alterar dados no gráfico da cópia. Nunca propaga erro.
Public Sub Core_NotifyCopiedChartOnce(co As ChartObject)
    Dim ws As Worksheet, n As Long, x As ChartObject, id As String, seen As Variant
    On Error Resume Next
    id = Mid$(co.Name, Len(ID_PREFIX) + 1)
    If Len(id) = 0 Then Exit Sub
    If Not mCopySeen Is Nothing Then
        seen = mCopySeen.Item(id)
        If Err.Number = 0 Then Exit Sub
        Err.Clear
    End If
    For Each ws In co.Parent.Parent.Worksheets
        Set x = Nothing
        Set x = ws.ChartObjects(co.Name)
        If Not x Is Nothing Then n = n + 1
        Err.Clear
    Next
    If n < 2 Then Exit Sub
    If mCopySeen Is Nothing Then Set mCopySeen = New Collection
    mCopySeen.Add id, id
    Core_Notify "Este gráfico do ThinQQ também existe em outra planilha (planilha copiada): a cópia continua ligada aos dados e às opções do gráfico original. Use 'Alterar dados' no gráfico da cópia para separá-lo.", vbExclamation
End Sub

' Avisa uma única vez por gráfico e por sessão que os dados passaram do máximo da escala fixa (Mesma escala): a escala fixa é imposta
' em silêncio e as barras ficam cortadas. Nunca propaga erro.
Private Sub NotifyScaleClippedOnce(cht As Chart)
    Dim id As String, seen As Variant
    On Error Resume Next
    id = Registry_IdFromChart(cht)
    If Len(id) = 0 Or Len(Registry_GetOption(cht, "vmax", "")) = 0 Then Exit Sub
    If Not mClipSeen Is Nothing Then
        seen = mClipSeen.Item(id)
        If Err.Number = 0 Then Exit Sub
        Err.Clear
    End If
    If Not Deco_ScaleClipped(cht) Then Exit Sub
    If mClipSeen Is Nothing Then Set mClipSeen = New Collection
    mClipSeen.Add id, id
    ' cascata comparativa com as barras dentro da escala: so a sala das setas passou do maximo (Mesma escala de novo daria a mesma uniao)
    If Registry_ChartType(cht) = "waterfallpair" And Not Deco_ScaleClipped(cht, True) Then
        Core_Notify "As setas de variação passam do máximo da escala fixa (Mesma escala) e ficam cortadas. Use 'Escala automática' ou desligue as setas de variação.", vbExclamation
    Else
        Core_Notify "Os dados deste gráfico passaram do máximo da escala fixa (Mesma escala) e as barras ficam cortadas. Use 'Mesma escala' de novo ou 'Escala automática'.", vbExclamation
    End If
End Sub

' Reposiciona rótulos e anotações. Chamado por eventos: nunca propaga erro (apenas registra).
' Gráficos de planilhas inativas são atualizados quando a planilha é ativada (Watchers_Rehook).
Public Sub Core_SafeRefresh(cht As Chart)
    On Error Resume Next
    NotifySourceMismatchOnce cht
    Deco_ApplyFixedScale cht                         ' mesma escala: os dados mudaram e o eixo automatico recalculou; reaplica a escala gravada
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Escala" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Break_RefreshAfterDataChange cht                 ' corte de eixo: maximo fixo, folga do total, rotulos dos segmentos e marcas
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Break" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Labels_RefreshPositions cht
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Labels" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    ' barras: a auxiliar XY usa o eixo X secundario (se o Excel o criou); mantem-no alinhado ao eixo de valores primario
    If Kinds_IsHorizontal(Registry_ChartType(cht)) Then Chart_SyncSecondaryX cht
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/SyncX" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Combo_RefreshAfterDataChange cht                 ' combinacao: escala do eixo secundario e formato dos rotulos da linha
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Combo" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Pie_RefreshAfterDataChange cht                   ' pizza com rotulos fora: o nome acompanha a fatia que passou a ter (ou deixou de ter) valor
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Pie" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Bfly_RefreshAfterDataChange cht                  ' borboleta: o vao central acompanha o maior nome e o tamanho do grafico
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Bfly" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Wfp_RefreshAfterDataChange cht                   ' cascata comparativa: meia-largura das barras, valor por ponto das setas, rotulos finos, nomes ou legenda
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Wfp" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    NotifyScaleClippedOnce cht                       ' depois da recalibracao das setas da comparativa (a sala delas entra na conferencia)
    Scatter_RefreshAfterDataChange cht               ' dispersao/bolhas: quadrantes nas medias atuais e tendencias fora da legenda
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Scatter" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Gantt_RefreshScale cht                           ' Gantt: a escala calculada no bloco (unidade e intervalo) vai para o eixo; antes dos colchetes
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Gantt" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Ann_Refresh cht
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Ann" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
End Sub

' Re-layout do gráfico ativo (ou de todos da planilha) e limpeza de blocos órfãos.
Public Sub Core_Refresh()
    On Error GoTo Fail
    Dim c As Chart, co As ChartObject, ws As Object, wb As Workbook
    Set c = Core_ActiveChart()
    If Not c Is Nothing Then
        Core_SafeRefresh c
        CalcSheet_Cleanup Registry_Workbook(c)
    Else
        If ActiveWorkbook Is Nothing Then Exit Sub
        Set wb = ActiveWorkbook
        Set ws = ActiveSheet
        If Not ws Is Nothing Then
            If TypeName(ws) = "Worksheet" Then
                For Each co In ws.ChartObjects
                    If Left$(co.Name, Len(ID_PREFIX)) = ID_PREFIX Then Core_SafeRefresh co.Chart
                Next
            End If
        End If
        CalcSheet_Cleanup wb
    End If
    Watchers_Rehook
    Exit Sub
Fail:
    Errors_Handle "Core_Refresh"
End Sub

' Troca a fonte de dados: pede a nova faixa e recria o gráfico no mesmo lugar, com as mesmas opções. Em build Test sai antes do InputBox
' (modal travaria o Excel invisivel); os testes chamam Core_RebuildFromRange.
Public Sub Core_ChangeSource()
    On Error GoTo Fail
    Dim c As Chart, rng As Range
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    If BUILD_CONFIG = "Test" Then Exit Sub                ' build Test: sem InputBox
    On Error Resume Next
    Set rng = Application.InputBox("Selecione a nova tabela (cabeçalho + categorias):", APP_NAME, Type:=8)
    On Error GoTo Fail
    If rng Is Nothing Then Exit Sub
    Core_RebuildFromRange c, rng
    Exit Sub
Fail:
    Errors_Handle "Core_ChangeSource"
End Sub

' Recria o gráfico c a partir da faixa rng, na mesma posição, com as mesmas opções e toggles.
' O novo gráfico é construído ANTES; o antigo só é removido se a construção der certo (senão fica intacto).
' Anotações (setas) não são preservadas. Retorna True se recriou.
Public Function Core_RebuildFromRange(c As Chart, rng As Range) As Boolean
    On Error GoTo Fail
    Dim co As ChartObject, oldCo As ChartObject, data As clsChartData, msg As String
    Dim kind As String, opts As String, ws As Worksheet, wb As Workbook, nc As Chart, prevSU As Boolean
    Dim l As Double, t As Double, w As Double, h As Double, keepKeys As Variant, keepVals(5) As String, ki As Long
    prevSU = Application.ScreenUpdating
    Set oldCo = c.Parent
    Set ws = oldCo.Parent: Set wb = ws.Parent
    ' O endereço externo das fórmulas não leva o nome da pasta de trabalho: só vale para a pasta do gráfico.
    If Not rng.Worksheet.Parent Is wb Then
        Core_Notify "Selecione uma tabela na mesma pasta de trabalho do gráfico.", vbExclamation
        Exit Function
    End If
    Set data = ChartData_FromRange(rng)
    msg = data.Validate()
    If Len(msg) > 0 Then Core_Notify msg, vbExclamation: Exit Function
    If ws.ProtectContents Then
        Core_Notify "A planilha está protegida. Desproteja a planilha para trocar a fonte de dados.", vbExclamation
        Exit Function
    End If
    kind = Registry_ChartType(c)
    l = oldCo.Left: t = oldCo.Top: w = oldCo.Width: h = oldCo.Height
    opts = CStr(Registry_Block(c).Cells(1, 4).Value)
    Break_TakeNotice                            ' descarta um aviso de corte que tenha sobrado
    Application.ScreenUpdating = False
    ' O novo gráfico nasce na planilha do gráfico antigo (ws), não na da nova fonte.
    Set co = Kinds_Insert(ws, data, kind)
    Set nc = co.Chart
    On Error GoTo FailNew
    ' escala e formato dos eixos da dispersao/bolhas: os que o grafico NOVO calculou para os dados novos (os do antigo fariam o refresh nao
    ' reconhecer a propria escala e congelar o eixo); o eixo que o usuario fixou a mao no antigo passa ao novo (Scatter_KeepUserAxes)
    keepKeys = Array("xscale", "yscale", "xfmt", "yfmt", "xsrc", "ysrc")
    For ki = 0 To 5
        keepVals(ki) = Registry_GetOption(nc, CStr(keepKeys(ki)), "")
    Next
    Registry_Block(nc).Cells(1, 4).Value = opts
    For ki = 0 To 5
        If Registry_GetOption(nc, CStr(keepKeys(ki)), "") <> keepVals(ki) Then Registry_SetOption nc, CStr(keepKeys(ki)), keepVals(ki)
    Next
    If Kinds_Family(kind) = "scatter" Then Scatter_KeepUserAxes c, nc
    Registry_SetOption nc, "n", CStr(data.CategoryCount)     ' as opções antigas trazem o n da tabela anterior
    If Len(Registry_GetOption(nc, "thinlbl", "")) > 0 Then Registry_SetOption nc, "thinlbl", ""      ' rotulos finos escondidos: pontos do grafico antigo
    If Len(Registry_GetOption(nc, "outlbl", "")) > 0 Then Registry_SetOption nc, "outlbl", ""        ' rotulos postos fora da barra fina (borboleta): idem
    If Len(Registry_GetOption(nc, "outpos", "")) > 0 Then Registry_SetOption nc, "outpos", ""        ' e as posicoes aplicadas a eles
    If Len(Registry_GetOption(nc, "brklbl", "")) > 0 Then Registry_SetOption nc, "brklbl", ""        ' rotulos ocultos pelo corte de eixo: idem
    ReapplyOptions nc                           ' formato, paleta, espaco, alternadores, escala fixa e legenda: as ESCOLHAS salvas do usuario
    RebuildTestHook "afterreapply"              ' (teste) falha depois de reaplicar, antes de trocar o grafico
    oldCo.Delete
    ' A partir daqui o gráfico antigo não existe mais: nada pode apagar o novo. Falhas só são registradas.
    On Error Resume Next
    CalcSheet_Cleanup wb
    co.Left = l: co.Top = t: co.Width = w: co.Height = h
    Chart_EnsureRoomAfterChange nc              ' a folga dos totais (barras) vale para o tamanho do momento: refaz no tamanho final
    Application.ScreenUpdating = prevSU
    Watchers_Rehook
    If Not ActiveSheet Is Nothing Then
        If ActiveSheet.Name <> ws.Name Then ws.Activate
    End If
    co.Select
    If Err.Number <> 0 Then Errors_Log "Core_RebuildFromRange/final" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Core_SafeRefresh nc
    NotifyPieIgnoredColumns kind, data
    NotifyMekkoNegatives kind, data
    msg = Break_TakeNotice()                    ' corte de eixo: aviso so depois de o grafico antigo ser trocado
    If Len(msg) > 0 Then Core_Notify msg, vbInformation
    Core_RebuildFromRange = True
    Exit Function
FailNew:
    ' Falha antes de remover o antigo: descarta o novo e mantém o antigo.
    Dim en As Long, ed As String, nid As String
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    nid = Mid$(co.Name, Len(ID_PREFIX) + 1)
    co.Delete
    CalcSheet_DeleteBlock wb, nid
    Break_TakeNotice                            ' o antigo fica (cortado): o aviso do novo e descartado
    Application.ScreenUpdating = prevSU
    On Error GoTo 0
    Errors_HandleErr "Core_RebuildFromRange", en, ed
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    Application.ScreenUpdating = prevSU
    If en >= vbObjectError + 6007 And en <= vbObjectError + 6099 Then
        ' Erro de validação dos dados da nova tabela (6007 cascata, 6008 combinação, ...): aviso, não é falha; o gráfico antigo fica.
        Core_Notify ed, vbExclamation
    Else
        Errors_HandleErr "Core_RebuildFromRange", en, ed
    End If
End Function

' Reaplica no gráfico NOVO as escolhas do usuário (as opções foram copiadas do gráfico antigo): o Insert criou o PADRÃO (rótulos ligados,
' indicador 100% ligado nos 100%, sem linha de valor, eixo automático, legenda à direita) e a opção salva é uma ESCOLHA, nunca o estado real:
' um indicador ou rótulo desligado pelo usuário continua desligado, a linha de valor ligada volta e a escala fixa (vmin/vmax, Mesma escala) é
' reaplicada ao eixo (os dados novos não a mudam: quem quer outra escala usa Mesma escala de novo). Só o que o tipo suporta (Kinds_Supports);
' valores ausentes: rótulos e indicador 100% ligados, os demais alternadores desligados. Ordem: formato, paleta, espaço, alternadores (rótulos
' antes da legenda: "fim da linha" depende deles), escala fixa e, por último, a legenda. Fica aqui a lista única do que "Alterar dados" reaplica.
' Propaga erros (Core_RebuildFromRange descarta o gráfico novo e mantém o antigo); a legenda só registra (ReapplyLegend).
Private Sub ReapplyOptions(nc As Chart)
    Dim kind As String, k As Variant, dflt As String
    kind = Registry_ChartType(nc)
    If Kinds_Supports(kind, "numfmt") Then ApplyOption nc, "numfmt", Registry_GetOption(nc, "numfmt", "int")
    ApplyOption nc, "palette", Registry_GetOption(nc, "palette", "padrao")
    If Kinds_Supports(kind, "gap") And Len(Registry_GetOption(nc, "gap", "")) > 0 Then ApplyOption nc, "gap", Registry_GetOption(nc, "gap", "")
    ' soma consistente ligada: as colunas Rnd_j antes dos rotulos, para a passada dos rotulos ja ligar a elas (a de roundsum entao nao religa)
    If Kinds_Supports(kind, "roundsum") And Registry_GetOption(nc, "roundsum", "0") = "1" Then Round_EnsureColumns nc
    For Each k In Array("totals", "labels", "connectors", "secondary", "ind100", "vline", "pieout", "trend", "quad", "roundsum", "pairdiff", "pairabs")      ' roundsum depois de labels: com os rotulos desligados so grava a opcao
        If Kinds_Supports(kind, CStr(k)) Then
            dflt = IIf(k = "labels" Or k = "ind100" Or k = "pairdiff", "1", "0")
            ApplyToggle nc, CStr(k), (Registry_GetOption(nc, CStr(k), dflt) = "1")      ' pizza: rotulos fora (pieout) tambem aplicados aqui
        End If
    Next
    Deco_ApplyFixedScale nc
    Break_TryReapply nc                         ' corte de eixo (vbreak): coluna LblPos, eixo, rotulos e marcas no grafico novo; nunca propaga erro
    ReapplyLegend nc
End Sub

' Reaplica a opção de legenda que o usuário tinha (copiada do gráfico antigo para o novo pelas opções): a opção salva é uma
' ESCOLHA do usuário, não o estado real. O Insert do novo gráfico criou a legenda padrão (à direita, ou nenhuma na cascata de
' uma série); aqui "none"/"lineend" voltam a valer. Tipo sem legenda possível (cascata de uma série, Gantt): a escolha salva
' fica como está (o gráfico não tem legenda; se a tabela voltar a ter várias séries, a escolha volta a valer). Sem opção salva
' (gráfico sem a chave), fica o estado que o Insert produziu e a opção passa a refleti-lo. Nunca propaga erro (registra).
Private Sub ReapplyLegend(nc As Chart)
    Dim mode As String
    On Error GoTo Fail
    mode = Registry_GetOption(nc, "legend", "")
    If Len(mode) = 0 Then
        If nc.HasLegend Then Registry_SetOption nc, "legend", "right"
        Exit Sub
    End If
    If Not Kinds_SupportsChart(nc, "legend") Then Exit Sub
    If mode = "lineend" And Not Kinds_Supports(Registry_ChartType(nc), "lineend") Then mode = "right"
    Deco_SetLegend nc, mode
    Exit Sub
Fail:
    Errors_Log "ReapplyLegend" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Substitui o arquivo do loader (registrado em PendingLoaderPath) e o recarrega.
' Executado via OnTime a partir do core, para que nenhum codigo do loader esteja na pilha.
' Seguranca: o caminho pendente vem do registro, entao o SHA-256 gravado pelo atualizador
' (PendingLoaderSha256) e recalculado aqui (Util_Sha256File, SHA-256 em VBA puro) sobre a copia <alvo>.new, ou seja, os bytes
' verificados sao os bytes instalados; divergencia => descarta o arquivo.
' Atomicidade: .new pronto e verificado ANTES de fechar o loader; depois de fechado, o alvo vira .bak (renomeacao:
' um arquivo aberto no Excel nao pode ser copiado) e o .new vira o alvo. Qualquer falha restaura o .bak, recarrega o
' loader antigo e deixa o pendente registrado.
' Fechar o loader dispara o Workbook_BeforeClose dele (Loader_Shutdown), que fecharia ESTE core (quem executa a
' troca). Por isso o registro recebe SwapInProgress (timestamp) antes de mexer no loader e o loader nao fecha o core
' enquanto a flag estiver recente (< 5 min). A flag e limpa em toda saida.
' Guarda de reentrada (o hash usa DoEvents; dois OnTime podem chegar). Nunca propaga erro e nunca usa MsgBox direto.
Public Sub Core_ApplyPendingLoader()
    If mApplyBusy Then Exit Sub
    mApplyBusy = True
    ApplyPendingLoaderImpl
    On Error Resume Next
    DeleteSetting "ThinQQ", "Updater", "SwapInProgress"
    mApplyBusy = False
End Sub

' Gancho de teste da guarda de reentrada.
Public Sub Core_TestSetApplyBusy(value As Boolean)
    If BUILD_CONFIG <> "Test" Then Exit Sub     ' gancho de teste: inerte em Release
    mApplyBusy = value
End Sub

' Gancho de teste: forca a familia de despacho de ApplyToggle/ApplyOption (simula uma divergencia entre a matriz de capacidades e o despacho:
' um tipo que a matriz aceita mas cuja familia nao tem Case). "" desfaz. Inerte em Release.
Public Sub Core_TestSetFamilyOverride(fam As String)
    On Error Resume Next
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestFamily = fam
End Sub

' Gancho de teste: as respostas dos dois InputBox do Arredondar intervalo (casas decimais e celula de destino) para a PROXIMA chamada de
' Core_RoundRange (uso unico); dest Nothing desarma. Inerte em Release.
Public Sub Core_TestSetRoundRangeAnswers(dec As Variant, dest As Range)
    On Error Resume Next
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestRoundDec = dec
    Set mTestRoundDest = dest
End Sub

' Gancho de teste: a resposta do InputBox do Corte de eixo para a PROXIMA chamada de Core_SetAxisBreak (uso unico; False = Cancelar; Empty
' desarma). Inerte em Release.
Public Sub Core_TestSetAxisBreakAnswer(answer As Variant)
    On Error Resume Next
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestBreakAnswer = answer
End Sub

Private Sub ApplyPendingLoaderImpl()
    Dim pending As String, target As String, expectSha As String, nw As String, bak As String
    Dim ai As AddIn, found As Boolean, started As Boolean, committed As Boolean, en As Long, ed As String, prevAlerts As Boolean
    On Error GoTo Fail
    pending = GetSetting("ThinQQ", "Updater", "PendingLoaderPath", "")
    target = mLoaderPath
    If Len(pending) = 0 Or Len(target) = 0 Then Exit Sub
    If Len(Dir(pending)) = 0 Then
        SaveSetting "ThinQQ", "Updater", "PendingLoaderPath", ""
        SaveSetting "ThinQQ", "Updater", "PendingLoaderSha256", ""
        Exit Sub
    End If
    expectSha = LCase$(GetSetting("ThinQQ", "Updater", "PendingLoaderSha256", ""))
    nw = target & ".new": bak = target & ".bak"
    If Len(Dir(nw)) > 0 Then Kill nw
    If Len(Dir(bak)) > 0 Then Kill bak
    FileCopy pending, nw
    If Len(expectSha) = 0 Or Util_Sha256File(nw) <> expectSha Then
        Errors_Log "Core_ApplyPendingLoader: SHA-256 do loader pendente nao confere; descartado: " & pending
        Kill nw
        Kill pending
        SaveSetting "ThinQQ", "Updater", "PendingLoaderPath", ""
        SaveSetting "ThinQQ", "Updater", "PendingLoaderSha256", ""
        Core_Notify "A atualização do carregador do ThinQQ foi descartada: o arquivo não passou na verificação de integridade.", vbExclamation, False
        Exit Sub
    End If
    SaveSetting "ThinQQ", "Updater", "SwapInProgress", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    started = True
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For Each ai In Application.AddIns
        If StrComp(ai.FullName, target, vbTextCompare) = 0 Then ai.Installed = False: found = True
    Next
    On Error Resume Next
    Application.Workbooks(Util_FileName(target)).Close SaveChanges:=False
    On Error GoTo Fail
    ' O arquivo do loader aberto no Excel fica travado (nem copia nem apaga): o backup e feito por renomeacao,
    ' depois que o loader foi fechado.
    If Len(Dir(target)) > 0 Then Name target As bak
    Name nw As target
    If found Then
        For Each ai In Application.AddIns
            If StrComp(ai.FullName, target, vbTextCompare) = 0 Then ai.Installed = True
        Next
    Else
        Application.Workbooks.Open target
    End If
    committed = True
    On Error Resume Next
    Application.DisplayAlerts = prevAlerts
    Kill bak
    Kill pending
    SaveSetting "ThinQQ", "Updater", "PendingLoaderPath", ""
    SaveSetting "ThinQQ", "Updater", "PendingLoaderSha256", ""
    DeleteSetting "ThinQQ", "Updater", "PendingLoaderFailNotified"
    Errors_Log "loader substituído: " & target
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    ' A restauracao fica em outro procedimento: dentro de um handler ativo o On Error Resume Next nao protege
    ' de forma confiavel (um novo erro escapava para o chamador).
    If started And Not committed Then RestoreLoaderFiles target, nw, bak, found, prevAlerts
    Errors_Log "Core_ApplyPendingLoader" & vbTab & en & vbTab & ed & IIf(started And Not committed, " (loader antigo restaurado)", "")
    NotifyPendingFailureThrottled
End Sub

' Aviso de falha da troca do loader no maximo 1x por 24 h (o pendente e tentado a cada abertura do Excel).
Private Sub NotifyPendingFailureThrottled()
    Dim last As String
    On Error Resume Next
    last = GetSetting("ThinQQ", "Updater", "PendingLoaderFailNotified", "")
    If Len(last) > 0 Then
        If IsDate(last) Then
            If Now - CDate(last) >= 0 And Now - CDate(last) < 1 Then Exit Sub
        End If
    End If
    SaveSetting "ThinQQ", "Updater", "PendingLoaderFailNotified", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Core_Notify "Não foi possível atualizar o carregador do ThinQQ. A versão anterior foi mantida.", vbExclamation, False
End Sub

' Desfaz a troca do loader: devolve o backup, apaga o .new e recarrega o loader antigo. Nunca propaga erro.
Private Sub RestoreLoaderFiles(target As String, nw As String, bak As String, wasAddin As Boolean, prevAlerts As Boolean)
    Dim ai As AddIn
    On Error Resume Next
    If Len(Dir(bak)) > 0 Then
        If Len(Dir(target)) > 0 Then Kill target
        Name bak As target
    End If
    If Len(Dir(nw)) > 0 Then Kill nw
    If wasAddin Then
        For Each ai In Application.AddIns
            If StrComp(ai.FullName, target, vbTextCompare) = 0 Then ai.Installed = True
        Next
    Else
        Application.Workbooks.Open target
    End If
    Application.DisplayAlerts = prevAlerts
End Sub

' Abre uma pasta nova com tabelas e gráficos de exemplo. Nunca propaga erro.
Public Sub Core_InsertSamples()
    Dim prevSU As Boolean
    On Error GoTo Fail
    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Samples_Create
    Application.ScreenUpdating = prevSU
    Watchers_Rehook
    Exit Sub
Fail:
    Application.ScreenUpdating = prevSU
    Errors_Handle "Core_InsertSamples"
End Sub

' Copiar para o PowerPoint: "image" (figura) ou "chart" (grafico do PowerPoint vinculado a pasta salva). Sem grafico ativo nao faz nada;
' a pasta nao salva e recusada ANTES de abrir o PowerPoint; PowerPoint ausente (429) vira aviso. Nunca propaga erro. Em falha desfaz o que
' criou (forma colada, apresentacao nova, PowerPoint aberto por nos). A colagem usa a area de transferencia do Windows.
Public Sub Core_CopyToPowerPoint(mode As String)
    On Error Resume Next
    CopyToPowerPointImpl mode
    Ppt_SetFailForTest ""                            ' gancho de teste: desarmado em toda saida (inclusive as antecipadas); o destino fixo e
                                                     ' desarmado pelo teste (Ppt_ResetTestHooks), pois vale para varias chamadas
End Sub

Private Sub CopyToPowerPointImpl(mode As String)
    On Error GoTo Fail
    Dim c As Chart, co As ChartObject, wb As Workbook, app As Object, slide As Object, shp As Object, msg As String, linked As Boolean
    Dim en As Long, ed As String, created As Boolean, newPres As Object, linkState As Long, newSlide As Object
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    If mode <> "image" And mode <> "chart" Then
        Errors_Log "Core_CopyToPowerPoint: modo desconhecido (ignorado)" & vbTab & mode
        Exit Sub
    End If
    linked = (mode = "chart")
    Set wb = Registry_Workbook(c)
    If linked Then
        msg = Ppt_LinkedPrecheck(wb)
        If Len(msg) > 0 Then Core_Notify msg, vbExclamation: Exit Sub
    End If
    Set app = Ppt_GetApp(True)
    created = Ppt_AppWasCreated()
    If app Is Nothing Then Core_Notify PPT_ERR_NOT_INSTALLED, vbExclamation: Exit Sub
    Set slide = Ppt_TargetSlide(app, newPres, newSlide)     ' o que ESTA chamada criou: apresentacao nova ou slide numa aberta vazia
    Ppt_TestHook "target"
    linkState = -1
    If linked Then
        Set co = c.Parent
        Set shp = Ppt_PasteChart(co, slide)
        On Error Resume Next: Application.CutCopyMode = False: Err.Clear: On Error GoTo Fail    ' a colagem ja deu certo: cancelar a copia nao pode desfazer nada
        linkState = Ppt_LinkState(shp)               ' o PowerPoint pode colar sem o vinculo: o aviso diz isso
        If linkState = 1 Then PptSyncAfterPaste shp, co     ' etiquetas e anotacoes no tamanho do slide (o ajuste ao slide estica as formas)
    Else
        Set shp = Ppt_PasteImage(c, slide)
    End If
    msg = Ppt_DoneText(CLng(slide.SlideIndex), CStr(slide.Parent.Name), linked, linkState)
    If linked And linkState <> 0 And Not wb.Saved Then msg = msg & " " & PPT_NOTE_DIRTY
    Ppt_TestHook "after"
    Core_Notify msg, vbInformation                   ' ANTES de ativar o PowerPoint: a caixa (Release) nao fica escondida atras dele
    If BUILD_CONFIG <> "Test" Then                   ' nos testes o PowerPoint do teste nao aparece nem rouba o foco
        On Error Resume Next
        app.Visible = -1
        app.Activate
        Err.Clear
        On Error GoTo Fail
    End If
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Application.CutCopyMode = False
    ' desfaz o que esta chamada criou, pelos objetos (nunca por indice: o usuario pode ter outras apresentacoes abertas): a forma colada,
    ' o slide que ela acrescentou a uma apresentacao do usuario, a apresentacao nova e o PowerPoint que ela abriu (so sem apresentacoes)
    If Not shp Is Nothing Then shp.Delete
    If Not newSlide Is Nothing Then
        If newPres Is Nothing Then newSlide.Delete
    End If
    If Not newPres Is Nothing Then
        newPres.Saved = -1
        newPres.Close
    End If
    If created And Not app Is Nothing Then
        If app.Presentations.Count = 0 Then app.Quit
    End If
    On Error GoTo 0
    If Ppt_IsKnownFailure(en) Then
        Core_Notify Ppt_ErrorText(en, ed), vbExclamation
    Else
        Errors_HandleErr "Core_CopyToPowerPoint(" & mode & ")", en, Ppt_ErrorText(en, ed)
    End If
End Sub

' Grafico vinculado recem-colado (Parte A): etiquetas TQQ_* (para a atualizacao ao reabrir) e as anotacoes recalculadas no tamanho do slide
' (medido: Ppt_PlaceShape estica as formas internas junto com o grafico). Falha aqui NAO desfaz a colagem: fica registrada (o botao
' "Atualizar graficos no PowerPoint" refaz). A impressao digital TQQ_FP so e gravada se a sincronizacao deu certo (senao a proxima
' atualizacao pularia o grafico como "sem mudanca" com as setas ainda esticadas).
Private Sub PptSyncAfterPaste(shp As Object, co As ChartObject)
    Dim ok As Boolean
    On Error Resume Next
    PptSync_TagShape shp, co
    If Err.Number <> 0 Then Errors_Log "CopyToPowerPoint/etiquetas" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    PptSync_SyncShape shp, co, False                 ' sem LinkFormat.Update: os dados acabaram de ser colados (o 1.3.0 tambem nao chamava)
    ok = (Err.Number = 0)
    If Not ok Then Errors_Log "CopyToPowerPoint/sincronizar" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    If ok Then shp.Tags.Add PPTSYNC_TAG_FP, PptSync_Fingerprint(co, shp)
    Err.Clear
End Sub

' Gancho de teste de Core_RebuildFromRange (Util_FailThinSyncForTest = "rebuild-<ponto>:1"): dispara UMA vez. Inerte fora do build Test.
Private Sub RebuildTestHook(site As String)
    If BUILD_CONFIG = "Test" Then
        If Util_FailThinSyncForTest = "rebuild-" & site & ":1" Then
            Util_FailThinSyncForTest = ""
            Err.Raise vbObjectError + 9991, "ThinQQ", "Falha simulada (teste) " & site
        End If
    End If
End Sub
