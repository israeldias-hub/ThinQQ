Attribute VB_Name = "modCoreApi"
Option Explicit
' API pública chamada pelo loader via Application.Run. Nunca propaga erro.
Private mLoaderPath As String
Private mApplyBusy As Boolean      ' guarda de reentrada de Core_ApplyPendingLoader
Private mLastNotice As String      ' última mensagem de Core_Notify (usada pelos testes)
Private mNoticeCount As Long        ' chamadas a Core_Notify desde Core_ClearLastNotice (testes)
Private mMismatchSeen As Collection ' ids de gráficos já avisados de tabela de origem com outro tamanho

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
        Core_Supports = Kinds_Supports(Registry_ChartType(c), feature)
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
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    Application.ScreenUpdating = prevSU
    If en = vbObjectError + 7200 Then
        ' Tipo válido ainda sem implementação nesta versão: aviso ao usuário, não é falha (nada foi criado).
        Core_Notify ed, vbInformation
    Else
        Errors_HandleErr "Core_Insert(" & kind & ")", en, ed
    End If
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
    Dim kind As String
    kind = Registry_ChartType(c)
    Select Case key
        Case "totals"
            If Kinds_Family(kind) <> "waterfall" Then Stacked_SetTotals c, value
        Case "labels"
            If Kinds_Family(kind) = "waterfall" Then Waterfall_SetLabels c, value Else Stacked_SetLabels c, value
        Case "connectors"
            If Kinds_Family(kind) = "waterfall" Then Ann_SetWaterfallConnectors c, value Else Stacked_SetConnectors c, value
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
    Registry_SetOption c, key, value
    Select Case key
        Case "numfmt"
            Style_ApplyNumberFormat c, Style_NumberFormat(value)
        Case "palette"
            If Kinds_Family(Registry_ChartType(c)) = "waterfall" Then Waterfall_ApplyColors c, value Else Style_ApplySeriesColors c, value
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
    Errors_Handle "Core_SetToggle(" & key & ")"
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
    ApplyOption c, key, value
    Core_SafeRefresh c
    Exit Sub
Fail:
    Errors_Handle "Core_SetOption(" & key & ")"
End Sub

Public Sub Core_AddAnnotation(kind As String, Optional fromIdx As Long = 0, Optional toIdx As Long = 0)
    On Error GoTo Fail
    Dim c As Chart, n As Long
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    n = Geom_ChartCategoryCount(c)
    If n < 2 Then Core_Notify "O gráfico precisa de pelo menos 2 categorias.": Exit Sub
    If fromIdx < 1 Or toIdx > n Or fromIdx >= toIdx Then
        If Not Core_AskCategories(c, n, fromIdx, toIdx) Then Exit Sub
    End If
    If kind = "cagr" Then Ann_AddCagr c, fromIdx, toIdx Else Ann_AddDifference c, fromIdx, toIdx
    Exit Sub
Fail:
    If Err.Number = vbObjectError + 7001 Then
        ' Seta (CAGR/diferença) em tipo que não a suporta (recusa da matriz de tipos): aviso ao usuário, não é falha.
        Core_Notify Err.Description, vbInformation
    Else
        Errors_Handle "Core_AddAnnotation(" & kind & ")"
    End If
End Sub

' Abre o formulário de categorias. Nunca abre em build de teste (modal travaria o Excel invisível).
Private Function Core_AskCategories(c As Chart, n As Long, ByRef fromIdx As Long, ByRef toIdx As Long) As Boolean
    Dim frm As frmAnnotation, cats() As String, body As Range, i As Long
    If BUILD_CONFIG = "Test" Then Exit Function
    Set body = Block_Body(Registry_Block(c))
    ReDim cats(1 To n)
    For i = 1 To n
        cats(i) = Util_CellLabel(body.Cells(i + 1, 1))
    Next
    Set frm = New frmAnnotation
    frm.Prepare cats, 1, n
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

' Reposiciona rótulos e anotações. Chamado por eventos: nunca propaga erro (apenas registra).
' Gráficos de planilhas inativas são atualizados quando a planilha é ativada (Watchers_Rehook).
Public Sub Core_SafeRefresh(cht As Chart)
    On Error Resume Next
    NotifySourceMismatchOnce cht
    Labels_RefreshPositions cht
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/Labels" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    ' barras: a auxiliar XY usa o eixo X secundario (se o Excel o criou); mantem-no alinhado ao eixo de valores primario
    If Kinds_IsHorizontal(Registry_ChartType(cht)) Then Chart_SyncSecondaryX cht
    If Err.Number <> 0 Then Errors_Log "Core_SafeRefresh/SyncX" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
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
    ApplyToggle nc, "labels", Registry_GetOption(nc, "labels", "1") = "1"
    ApplyToggle nc, "connectors", Registry_GetOption(nc, "connectors", "0") = "1"
    oldCo.Delete
    ' A partir daqui o gráfico antigo não existe mais: nada pode apagar o novo. Falhas só são registradas.
    On Error Resume Next
    CalcSheet_Cleanup wb
    co.Left = l: co.Top = t: co.Width = w: co.Height = h
    Application.ScreenUpdating = prevSU
    Watchers_Rehook
    If Not ActiveSheet Is Nothing Then
        If ActiveSheet.Name <> ws.Name Then ws.Activate
    End If
    co.Select
    If Err.Number <> 0 Then Errors_Log "Core_RebuildFromRange/final" & vbTab & Err.Number & vbTab & Err.Description: Err.Clear
    Core_SafeRefresh nc
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
    Application.ScreenUpdating = prevSU
    Errors_Handle "Core_RebuildFromRange"
End Function

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
