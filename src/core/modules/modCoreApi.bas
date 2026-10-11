Attribute VB_Name = "modCoreApi"
Option Explicit
' API pública chamada pelo loader via Application.Run. Nunca propaga erro.
Private mLoaderPath As String
Private mApplyBusy As Boolean      ' guarda de reentrada de Core_ApplyPendingLoader
Private mLastNotice As String      ' última mensagem de Core_Notify (usada pelos testes)
Private mNoticeCount As Long        ' chamadas a Core_Notify desde Core_ClearLastNotice (testes)
Private mMismatchSeen As Collection ' ids de gráficos já avisados de tabela de origem com outro tamanho
Private mClipSeen As Collection     ' ids de gráficos já avisados de dados acima da escala fixa (Mesma escala)

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
        Core_Supports = True
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
    Application.ScreenUpdating = prevSU
    Watchers_Rehook
    co.Select
    Core_SafeRefresh co.Chart
    NotifyPieIgnoredColumns kind, data
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

' Pizza e rosca mostram uma serie so: a primeira coluna de valores. Com mais colunas na tabela as demais sao ignoradas e o usuario e avisado
' (informativo: o grafico ja foi criado). Nunca propaga erro.
Private Sub NotifyPieIgnoredColumns(kind As String, data As clsChartData)
    On Error Resume Next
    If Kinds_Family(kind) <> "pie" Then Exit Sub
    If data.SeriesCount < 2 Then Exit Sub
    Core_Notify "A " & IIf(kind = "doughnut", "rosca", "pizza") & " mostra só a primeira coluna de valores (" & data.SeriesName(1) & "). As outras " & (data.SeriesCount - 1) & " coluna(s) da tabela foram ignoradas.", vbInformation
End Sub

Public Function Core_GetToggle(key As String) As Boolean
    Dim c As Chart
    On Error Resume Next
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Function
    Core_GetToggle = (Registry_GetOption(c, key, "0") = "1")
End Function

' Aplica o toggle no gráfico informado (sem depender da seleção). Propaga erros.
Private Sub ApplyToggle(c As Chart, key As String, value As Boolean)
    Dim kind As String, fam As String
    kind = Registry_ChartType(c)
    fam = Kinds_Family(kind)
    If fam = "pie" Then
        ' Pizza e rosca: so os rotulos (nativos) e os rotulos fora (a rosca recusa: 7002). Totais, conectores e eixo secundario nao existem
        ' e nao gravam nada (Core_RebuildFromRange chama todos); linha de valor e indicador 100% seguem para a recusa propria (7002).
        Select Case key
            Case "labels": Pie_SetLabels c, value: Exit Sub
            Case "pieout": Pie_SetOutside c, value: Exit Sub
            Case "ind100", "vline"
            Case Else: Exit Sub
        End Select
    End If
    If fam = "gantt" Then
        ' Gantt: so os rotulos a direita das barras (Gantt_SetLabels). Totais, conectores e eixo secundario nao existem e nao gravam nada
        ' (Core_RebuildFromRange chama todos); linha de valor, indicador 100% e rotulos fora seguem para a recusa propria (7002).
        Select Case key
            Case "labels": Gantt_SetLabels c, value: Exit Sub
            Case "ind100", "vline", "pieout"
            Case Else: Exit Sub
        End Select
    End If
    Select Case key
        Case "totals"
            If fam = "area" Then
                Area_SetTotals c, value
            ElseIf fam = "mekko" Then
                Mekko_SetTotals c, value
            ElseIf fam <> "waterfall" Then
                Stacked_SetTotals c, value              ' inclui a combinacao (so os totais das colunas)
            End If
        Case "labels"
            If fam = "waterfall" Then
                Waterfall_SetLabels c, value
            ElseIf fam = "area" Then
                Area_SetLabels c, value
            ElseIf fam = "mekko" Then
                Mekko_SetLabels c, value
            Else
                Stacked_SetLabels c, value              ' inclui a combinacao
            End If
        Case "connectors"
            If fam = "waterfall" Then Ann_SetWaterfallConnectors c, value Else Stacked_SetConnectors c, value
        Case "secondary"
            Combo_SetSecondary c, value                 ' so a combinacao (nos demais tipos nao faz nada nem grava a opcao)
        Case "pieout"
            Pie_SetOutside c, value                     ' recusa (7002) fora da pizza (a pizza e a rosca saem acima): Core_SetToggle mostra o aviso
        Case "ind100"
            Deco_SetInd100 c, value                     ' recusa (7002) fora dos 100% e Mekko (%): Core_SetToggle mostra o aviso
        Case "vline"
            Deco_SetValueLine c, value                  ' recusa (7002) nos tipos sem linha de valor ou em grafico antigo
    End Select
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
                Case "right", "none", "lineend"
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

' Aplica a opção no gráfico informado (sem depender da seleção). Valida antes de persistir. Propaga erros.
Private Sub ApplyOption(c As Chart, key As String, value As String)
    Dim msg As String
    msg = ValidateOption(key, value)
    If Len(msg) > 0 Then Err.Raise vbObjectError + 7100, "Core", msg
    If key = "legend" Then
        ' A legenda valida, aplica e SO ENTAO grava (Deco_SetLegend): uma recusa (7002) nao pode deixar legend=lineend gravado.
        ' As demais opcoes gravam antes de despachar porque o despacho le a opcao do registro (ex.: numfmt, em Labels_SegmentFormat).
        Deco_SetLegend c, value
        Exit Sub
    End If
    Registry_SetOption c, key, value
    Select Case key
        Case "numfmt"
            Style_ApplyNumberFormat c, Style_NumberFormat(value)
            Chart_EnsureRoomAfterChange c               ' rotulos de total mais largos: refaz a folga dos totais e das setas
        Case "palette"
            Select Case Kinds_Family(Registry_ChartType(c))
                Case "waterfall": Waterfall_ApplyColors c, value
                Case "area": Area_ApplyColors c, value
                Case "pie": Pie_ApplyColors c, value
                Case "mekko": Mekko_ApplyColors c, value        ' separadores de 1,5 pt e contraste dos rotulos (por ponto)
                Case "gantt": Gantt_ApplyColors c, value        ' barras pal(1) e marcos pal(0) (as series do Gantt sao todas tqq_*)
                Case Else: Style_ApplySeriesColors c, value
            End Select
        Case "gap"
            c.ChartGroups(1).GapWidth = CLng(Val(value))
    End Select
End Sub

Public Sub Core_SetToggle(key As String, value As Boolean)
    On Error GoTo Fail
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    ApplyToggle c, key, value
    Core_SafeRefresh c
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
    If key = "gap" And Registry_ChartType(c) = "line" Then
        Core_Notify "Espaço entre colunas não se aplica a gráficos de linha.", vbInformation
        Exit Sub
    End If
    If key = "gap" And Not Kinds_Supports(Registry_ChartType(c), "gap") Then
        Core_Notify "Espaço entre colunas não se aplica a este tipo de gráfico.", vbInformation
        Exit Sub
    End If
    ApplyOption c, key, value
    Core_SafeRefresh c
    Exit Sub
Fail:
    If Err.Number = vbObjectError + 7002 Then
        Core_Notify Err.Description, vbInformation       ' recusa (ex.: "fim da linha" fora de linha/combinacao): aviso, nao e falha
    Else
        Errors_Handle "Core_SetOption(" & key & ")"
    End If
End Sub

' Setas (cagr, diff) e colchete (bracket, so no Gantt). O colchete vale sobre UMA atividade (de = ate): exige n >= 1 e so pede as atividades
' quando de/ate nao formam um intervalo valido (de > ate ou fora de 1..n); recusado (7001) antes do formulario fora do Gantt. As setas seguem
' exigindo n >= 2 e de < ate.
Public Sub Core_AddAnnotation(kind As String, Optional fromIdx As Long = 0, Optional toIdx As Long = 0)
    On Error GoTo Fail
    Dim c As Chart, n As Long
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    n = Geom_ChartCategoryCount(c)
    If kind = "bracket" Then
        If Not Kinds_SupportsChart(c, "bracket") Then Err.Raise vbObjectError + 7001, "Core_AddAnnotation", ANN_BRACKET_REFUSAL
        If n < 1 Then Core_Notify "O gráfico precisa de pelo menos 1 atividade.": Exit Sub
        If fromIdx < 1 Or toIdx > n Or fromIdx > toIdx Then
            If Not Core_AskCategories(c, n, fromIdx, toIdx, True) Then Exit Sub
        End If
        Ann_AddBracket c, fromIdx, toIdx
        Exit Sub
    End If
    If n < 2 Then Core_Notify "O gráfico precisa de pelo menos 2 categorias.": Exit Sub
    If fromIdx < 1 Or toIdx > n Or fromIdx >= toIdx Then
        If Not Core_AskCategories(c, n, fromIdx, toIdx) Then Exit Sub
    End If
    If kind = "cagr" Then Ann_AddCagr c, fromIdx, toIdx Else Ann_AddDifference c, fromIdx, toIdx
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
    Dim coll As Collection, n As Long, skipped As Long, msg As String
    Set coll = Core_SelectedCharts()
    If coll.Count < 2 Then Set coll = SheetThinQQCharts()
    If coll.Count < 2 Then
        Core_Notify "Mesma escala: selecione dois ou mais gráficos do ThinQQ (ou nenhum, para usar todos os gráficos da planilha).", vbInformation
        Exit Sub
    End If
    n = Deco_SameScale(coll, skipped)
    If n < 2 Then
        Core_Notify "Mesma escala: são necessários dois ou mais gráficos com escala livre (100%, Mekko (%), pizza e Gantt não entram).", vbInformation
        Exit Sub
    End If
    msg = "Mesma escala aplicada a " & n & " gráfico(s)."
    If skipped > 0 Then msg = msg & " " & skipped & " ignorado(s): tipo sem escala livre (100%, Mekko (%), pizza, Gantt)."
    Core_Notify msg, vbInformation
    Exit Sub
Fail:
    Errors_Handle "Core_SameScale"
End Sub

' Escala automatica: devolve ao automatico o eixo de valores do grafico ThinQQ ativo ou dos selecionados (sem nenhum, de todos os da
' planilha ativa, como a Mesma escala) e limpa a escala gravada. Informa quantos graficos foram restaurados.
Public Sub Core_ResetScale()
    On Error GoTo Fail
    Dim coll As Collection, c As Variant, cht As Chart, n As Long
    Set coll = Core_SelectedCharts()
    If coll.Count = 0 Then Set coll = SheetThinQQCharts()
    If coll.Count = 0 Then
        Core_Notify "Escala automática: selecione um gráfico do ThinQQ (ou nenhum, para usar todos os gráficos da planilha).", vbInformation
        Exit Sub
    End If
    For Each c In coll
        Set cht = c
        If Kinds_Supports(Registry_ChartType(cht), "samescale") Then
            Deco_ResetScale cht
            n = n + 1
        End If
    Next
    If n = 0 Then
        Core_Notify "Escala automática: o gráfico não tem escala livre (100%, Mekko (%), pizza, Gantt).", vbInformation
    Else
        Core_Notify "Escala automática restaurada em " & n & " gráfico(s).", vbInformation
    End If
    Exit Sub
Fail:
    Errors_Handle "Core_ResetScale"
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
    Core_Notify "Os dados deste gráfico passaram do máximo da escala fixa (Mesma escala) e as barras ficam cortadas. Use 'Mesma escala' de novo ou 'Escala automática'.", vbExclamation
End Sub

' Reposiciona rótulos e anotações. Chamado por eventos: nunca propaga erro (apenas registra).
' Gráficos de planilhas inativas são atualizados quando a planilha é ativada (Watchers_Rehook).
Public Sub Core_SafeRefresh(cht As Chart)
    On Error Resume Next
    NotifySourceMismatchOnce cht
    Deco_ApplyFixedScale cht                         ' mesma escala: os dados mudaram e o eixo automatico recalculou; reaplica a escala gravada
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Escala" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    NotifyScaleClippedOnce cht
    Labels_RefreshPositions cht
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Labels" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    ' barras: a auxiliar XY usa o eixo X secundario (se o Excel o criou); mantem-no alinhado ao eixo de valores primario
    If Kinds_IsHorizontal(Registry_ChartType(cht)) Then Chart_SyncSecondaryX cht
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/SyncX" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Combo_RefreshAfterDataChange cht                 ' combinacao: escala do eixo secundario e formato dos rotulos da linha
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Combo" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Pie_RefreshAfterDataChange cht                   ' pizza com rotulos fora: o nome acompanha a fatia que passou a ter (ou deixou de ter) valor
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Pie" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
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

' Troca a fonte de dados: pede a nova faixa e recria o gráfico no mesmo lugar, com as mesmas opções.
Public Sub Core_ChangeSource()
    On Error GoTo Fail
    Dim c As Chart, rng As Range
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
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
    Dim l As Double, t As Double, w As Double, h As Double
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
    Application.ScreenUpdating = False
    ' O novo gráfico nasce na planilha do gráfico antigo (ws), não na da nova fonte.
    Set co = Kinds_Insert(ws, data, kind)
    Set nc = co.Chart
    On Error GoTo FailNew
    Registry_Block(nc).Cells(1, 4).Value = opts
    Registry_SetOption nc, "n", CStr(data.CategoryCount)     ' as opções antigas trazem o n da tabela anterior
    ApplyOption nc, "numfmt", Registry_GetOption(nc, "numfmt", "int")
    ApplyOption nc, "palette", Registry_GetOption(nc, "palette", "padrao")
    If Len(Registry_GetOption(nc, "gap", "")) > 0 Then ApplyOption nc, "gap", Registry_GetOption(nc, "gap", "")
    ApplyToggle nc, "totals", Registry_GetOption(nc, "totals", "0") = "1"
    ApplyToggle nc, "labels", Registry_GetOption(nc, "labels", "1") = "1"      ' pizza: aplica tambem os rotulos fora (pieout, copiado nas opcoes)
    ApplyToggle nc, "connectors", Registry_GetOption(nc, "connectors", "0") = "1"
    ApplyToggle nc, "secondary", Registry_GetOption(nc, "secondary", "0") = "1"      ' so a combinacao: a escolha do usuario, nao o padrao pelos dados novos
    ReapplySavedDecorations nc                  ' linha de valor, indicador 100% e mesma escala: a escolha salva do usuario, nao o padrao do Insert
    ReapplySavedLegend nc                       ' depois dos toggles: "fim da linha" depende dos rotulos ligados
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

' Reaplica as decoracoes que o usuario tinha (as opcoes foram copiadas do grafico antigo para o novo): o Insert do novo grafico criou o
' PADRAO (indicador 100% ligado nos 100%, sem linha de valor, eixo automatico) e a opcao salva e uma ESCOLHA, nao o estado real: um
' indicador desligado pelo usuario continua desligado, a linha de valor ligada volta e a escala fixa (vmin/vmax, Mesma escala) e reaplicada
' ao eixo (os dados novos nao a mudam: quem quer outra escala usa Mesma escala de novo). Opcao sem a chave: o padrao (indicador ligado,
' linha de valor desligada). So nos tipos com o recurso. Propaga erros (Core_RebuildFromRange descarta o grafico novo e mantem o antigo).
Private Sub ReapplySavedDecorations(nc As Chart)
    Dim kind As String
    kind = Registry_ChartType(nc)
    If Kinds_Supports(kind, "ind100") Then Deco_SetInd100 nc, (Registry_GetOption(nc, "ind100", "1") <> "0")
    If Kinds_Supports(kind, "vline") Then Deco_SetValueLine nc, (Registry_GetOption(nc, "vline", "0") = "1")
    Deco_ApplyFixedScale nc
End Sub

' Reaplica a opção de legenda que o usuário tinha (copiada do gráfico antigo para o novo pelas opções): a opção salva é uma
' ESCOLHA do usuário, não o estado real. O Insert do novo gráfico criou a legenda padrão (à direita, ou nenhuma na cascata de
' uma série); aqui "none"/"lineend" voltam a valer. Tipo sem legenda possível (cascata de uma série, Gantt): a escolha salva
' fica como está (o gráfico não tem legenda; se a tabela voltar a ter várias séries, a escolha volta a valer). Sem opção salva
' (gráfico sem a chave), fica o estado que o Insert produziu e a opção passa a refleti-lo. Nunca propaga erro (registra).
Private Sub ReapplySavedLegend(nc As Chart)
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
    Errors_Log "ReapplySavedLegend" & vbTab & Err.Number & vbTab & Err.Description
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
