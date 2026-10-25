Attribute VB_Name = "modRibbon"
Option Explicit
Private mRibbon As IRibbonUI
Private mHasChart As Boolean        ' cache de Core_HasChart por ciclo de Invalidate
Private mHasChartValid As Boolean
Private mSupports As Object         ' cache de Core_Supports por ciclo de Invalidate: Scripting.Dictionary tag -> Boolean
Private mSupportsMissing As Boolean ' o core nao tem Core_Supports (versao 1.0.0) ou nao esta carregado: vale ate o fim do ciclo
Private mStyleNames As Variant                 ' nomes dos estilos do ciclo de Invalidate (Rx_StyleNamesFor)
Private mStyleLegacy As Boolean                ' o core nao tem Core_StyleNames (anterior a 1.12.0): a lista sao as tres paletas
Private mStyleShown As Long                    ' quantos itens a lista Estilo exibiu (Rx_StyleCount): o clique confere
Private mSupportsProc As String     ' gancho de teste (Rx_TestSetSupportsProc); vazio = "Core_Supports"
Private mTestProcSuffix As String   ' gancho de teste (Rx_TestSetFeatureProcSuffix): simula um core sem os procedimentos novos
Public Const RX_MSG_UPDATE As String = "Atualize o ThinQQ para usar este recurso."
Public Const RX_MSG_STYLE_LIST_CHANGED As String = "A lista de estilos mudou. Abra a lista Estilo e escolha de novo."
Public Const TUTORIAL_URL As String = "https://israeldias-hub.github.io/ThinQQ/tutorial/tutorial.html"

Public Sub Rx_OnLoad(ribbon As IRibbonUI)
    On Error Resume Next
    Set mRibbon = ribbon
    ResetRibbonCaches
End Sub

' Zera os caches do ciclo de Invalidate (Core_HasChart e Core_Supports por tag). Cada ciclo le o tipo do grafico selecionado de novo.
Private Sub ResetRibbonCaches()
    mHasChartValid = False
    mSupportsMissing = False
    Set mSupports = Nothing
    mStyleNames = Empty
End Sub

' Gancho de teste: troca o nome do procedimento de capacidade (simula um core 1.0.0, sem Core_Supports). So existe no build de Teste.
Public Sub Rx_TestSetSupportsProc(procName As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mSupportsProc = procName
    ResetRibbonCaches
End Sub

' Gancho de teste: sufixo acrescentado ao nome dos procedimentos dos recursos sempre ativos (Arredondar intervalo, PowerPoint, Core_PptSyncOpen) para simular
' um core anterior que nao os tem ("" desarma). So existe no build de Teste.
Public Sub Rx_TestSetFeatureProcSuffix(suffix As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestProcSuffix = suffix
End Sub

' Recurso sempre ativo na faixa (sem consulta a matriz): com um core anterior carregado o procedimento nao existe e Loader_RunCore so
' registraria; avisa que e preciso atualizar (como Rx_OnSamples).
Private Sub RunCoreFeature(procName As String, Optional arg As Variant)
    Dim p As String
    p = procName
    If BUILD_CONFIG = "Test" Then p = p & mTestProcSuffix
    If IsMissing(arg) Then
        Loader_RunCore p
    ElseIf IsArray(arg) Then
        Loader_RunCore p, arg(0), arg(1)            ' Array(a, b): dois argumentos (as opcoes da atualizacao automatica)
    Else
        Loader_RunCore p, arg
    End If
    If Loader_RunCoreMissed() Then RxMsg RX_MSG_UPDATE, vbInformation
End Sub

' Unico ponto de MsgBox do modulo: em build de Teste apenas registra (nunca bloqueia).
Private Sub RxMsg(prompt As String, Optional style As Long = vbInformation)
    Loader_LastMsg = prompt                         ' a ultima mensagem da faixa tambem (lida pelos testes)
    If BUILD_CONFIG = "Test" Then
        Loader_Log "RxMsg (teste): " & prompt
    Else
        MsgBox prompt, style, APP_NAME
    End If
End Sub

' Cache de Core_HasChart: uma chamada ao core por ciclo de Invalidate.
Private Function HasChartCached() As Boolean
    If Not mHasChartValid Then
        mHasChart = False
        If Not Loader_CoreWorkbook() Is Nothing Then mHasChart = CBool(Loader_RunCore("Core_HasChart"))
        mHasChartValid = True
    End If
    HasChartCached = mHasChart
End Function

' Capacidade do grafico selecionado para a tag (matriz do core: Core_Supports). O resultado fica num Dictionary por ciclo de Invalidate:
' a faixa chama Rx_GetEnabled uma vez por controle e o core so e consultado uma vez por tag. A matriz depende do TIPO do grafico (e, na
' cascata, do numero de series): o cache nasce vazio a cada Loader_InvalidateRibbon, que o core dispara em toda mudanca de selecao
' (Watchers_NotifySelection) e que os comandos que trocam o grafico (inserir, dados, opcoes) tambem chamam. Core sem Core_Supports
' (Loader_RunCore devolve Empty): cai no comportamento antigo, "ha grafico selecionado" (HasChartCached), e nao insiste no resto do ciclo.
Private Function SupportsCached(tag As String) As Boolean
    Dim v As Variant, proc As String
    If mSupports Is Nothing Then Set mSupports = CreateObject("Scripting.Dictionary")
    If mSupports.Exists(tag) Then SupportsCached = CBool(mSupports(tag)): Exit Function
    v = Empty
    If Not mSupportsMissing Then
        proc = mSupportsProc
        If Len(proc) = 0 Then proc = "Core_Supports"
        v = Loader_RunCore(proc, tag)
        If IsEmpty(v) Then mSupportsMissing = True
    End If
    If IsEmpty(v) Then SupportsCached = HasChartCached() Else SupportsCached = CBool(v)
    mSupports.Add tag, SupportsCached
End Function

' Habilitacao de um controle pela tag: com tag, a capacidade do grafico selecionado; sem tag (Remover setas, Alterar dados, o menu Setas),
' basta haver um grafico ThinQQ selecionado. Nunca propaga erro (False em duvida).
Public Function Rx_IsEnabledFor(tag As String) As Boolean
    On Error Resume Next
    If Len(tag) > 0 Then Rx_IsEnabledFor = SupportsCached(tag) Else Rx_IsEnabledFor = HasChartCached()
    If Err.Number <> 0 Then Rx_IsEnabledFor = False
End Function

Public Sub Loader_InvalidateRibbon()
    On Error Resume Next
    ResetRibbonCaches
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

Private Function CoreOrWarn() As Boolean
    If Loader_CoreWorkbook() Is Nothing Then
        RxMsg "O módulo principal do ThinQQ não está carregado. Use 'Verificar atualizações' ou reinstale.", vbExclamation
        Exit Function
    End If
    CoreOrWarn = True
End Function

' Inserir (menus Colunas, Barras, Cascata, Linha e area, Mais tipos): so no Excel. O caminho para o slide e a galeria "Vincular ao
' PowerPoint" (Rx_OnLinkToPpt), desde a 1.13.0.
Public Sub Rx_OnInsert(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_Insert", control.Tag
    Loader_InvalidateRibbon                 ' o grafico novo fica selecionado: o tipo mudou
End Sub

' Galeria "Vincular ao PowerPoint" (1.13.0): o tipo vem do tag. Parte testavel em Rx_LinkProcFor.
Public Function Rx_LinkProcFor(tag As String) As String
    If tag = "table" Then
        Rx_LinkProcFor = "Core_PptInsertTable"
    ElseIf tag = "text" Then
        Rx_LinkProcFor = "Core_PptLinkTextFromPpt"
    ElseIf Len(tag) > 0 And InStr(tag, ",") = 0 And tag Like "[a-z]*" And Not tag Like "*[!a-z0-9]*" Then
        Rx_LinkProcFor = "Core_InsertToPowerPoint"
    End If
End Function

Public Sub Rx_OnLinkToPpt(control As IRibbonControl)
    Dim p As String
    On Error Resume Next
    p = Rx_LinkProcFor(CStr(control.Tag))
    If Len(p) = 0 Then Exit Sub
    If CoreOrWarn() Then
        If p = "Core_InsertToPowerPoint" Then RunCoreFeature p, CStr(control.Tag) Else RunCoreFeature p
    End If
    Loader_InvalidateRibbon
End Sub
' Mostrar/ocultar molduras (1.7.0): todas as molduras da pasta ativa. Sempre ativo; core anterior: aviso de atualizar.
Public Sub Rx_OnPptFrames(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_ToggleLinkFrames"
End Sub

Public Sub Rx_GetEnabled(control As IRibbonControl, ByRef enabled)
    On Error Resume Next
    enabled = False
    enabled = Rx_IsEnabledFor(control.Tag)
End Sub

Public Sub Rx_OnToggle(control As IRibbonControl, pressed As Boolean)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_SetToggle", control.Tag, pressed
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_GetPressed(control As IRibbonControl, ByRef pressed)
    On Error Resume Next
    pressed = False
    If Not Loader_CoreWorkbook() Is Nothing Then pressed = CBool(Loader_RunCore("Core_GetToggle", control.Tag))
End Sub

Public Sub Rx_OnDropdown(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_SetOption", control.Tag, Rx_OptionValueFor(control.Tag, selectedId)
    Loader_InvalidateRibbon
End Sub

' Valor da opcao que o core recebe para o item escolhido. Os itens do dropdown Espaco se chamam "gap30".."gap120" (um id de XML nao pode
' comecar por digito: com "30" o Excel descarta a aba inteira, em silencio); os demais dropdowns usam o proprio id.
Public Function Rx_OptionValueFor(tag As String, itemId As String) As String
    If tag = "gap" And Left$(itemId, 3) = "gap" Then Rx_OptionValueFor = Mid$(itemId, 4) Else Rx_OptionValueFor = itemId
End Function

' Item selecionado do dropdown da tag: a opcao gravada no grafico ou, sem opcao (grafico antigo, core sem a opcao), o padrao do item.
Public Function Rx_SelectedIdFor(tag As String) As String
    Dim v As Variant, s As String
    On Error Resume Next
    v = Loader_RunCore("Core_GetOption", tag)
    If Not IsEmpty(v) Then s = CStr(v)
    If Len(s) = 0 Then
        Select Case tag
            Case "numfmt": s = "int"
            Case "gap": s = "50"
            Case "legend": s = "right"
        End Select
    End If
    If tag = "gap" And Len(s) > 0 Then s = "gap" & s       ' id do item do dropdown Espaco (ver Rx_OptionValueFor)
    Rx_SelectedIdFor = s
End Function

Public Sub Rx_GetSelectedId(control As IRibbonControl, ByRef itemId)
    On Error Resume Next
    itemId = Rx_SelectedIdFor(control.Tag)
End Sub

' ===== Estilos (1.12.0) =====

' Lista de nomes a partir do texto do core ("a|b|c"); vazio (core anterior, sem Core_StyleNames, ou ausente) = as tres paletas de antes, que
' Rx_OnStyle aplica pela opcao palette (revisao final 1.12.0, M8). Pura.
Public Function Rx_StyleNamesFor(raw As String) As Variant
    If Len(raw) = 0 Then Rx_StyleNamesFor = Array("Padrão ThinQQ", "Tons de cinza", "Cores do tema") Else Rx_StyleNamesFor = Split(raw, "|")
End Function

' Paleta (opcao palette do core anterior) do item index da lista de reserva; "" fora dela. Pura.
Public Function Rx_LegacyPaletteFor(index As Long) As String
    Select Case index
        Case 0: Rx_LegacyPaletteFor = "padrao"
        Case 1: Rx_LegacyPaletteFor = "cinza"
        Case 2: Rx_LegacyPaletteFor = "tema"
    End Select
End Function

Private Function StyleNames() As Variant
    Dim v As Variant, p As String
    If IsEmpty(mStyleNames) Then
        mStyleLegacy = False
        If Not Loader_CoreWorkbook() Is Nothing Then
            p = "Core_StyleNames"
            If BUILD_CONFIG = "Test" Then p = p & mTestProcSuffix      ' gancho: simula um core anterior a 1.12.0
            v = Loader_RunCore(p)
            mStyleLegacy = Loader_RunCoreMissed()
        End If
        mStyleNames = Rx_StyleNamesFor(CStr(v))
    End If
    StyleNames = mStyleNames
End Function

Public Sub Rx_StyleCount(control As IRibbonControl, ByRef count)
    On Error Resume Next
    count = 0
    count = UBound(StyleNames()) + 1
    mStyleShown = count
End Sub

Public Sub Rx_StyleLabel(control As IRibbonControl, index As Integer, ByRef label)
    On Error Resume Next
    label = CStr(StyleNames()(index))
End Sub

Public Sub Rx_StyleId(control As IRibbonControl, index As Integer, ByRef id)
    On Error Resume Next
    id = "st" & index
End Sub

Public Sub Rx_GetStyleSel(control As IRibbonControl, ByRef itemId)
    Dim n As String, a As Variant, i As Long
    On Error Resume Next
    itemId = "st0"
    If Loader_CoreWorkbook() Is Nothing Then Exit Sub
    a = StyleNames()
    If mStyleLegacy Then
        ' core anterior: o item da paleta gravada no grafico
        Select Case CStr(Loader_RunCore("Core_GetOption", "palette"))
            Case "cinza": itemId = "st1"
            Case "tema": itemId = "st2"
        End Select
        Exit Sub
    End If
    n = CStr(Loader_RunCore("Core_GetStyleName"))
    For i = 0 To UBound(a)
        If StrComp(CStr(a(i)), n, vbTextCompare) = 0 Then
            itemId = "st" & i
            Exit Sub
        End If
    Next
End Sub

' O indice vem da lista EXIBIDA: se o cache foi refeito com outro tamanho entre a exibicao e o clique (um estilo salvo na janela), nada e
' aplicado e a lista e refeita (revisao final 1.12.0, M7). Core anterior: a paleta pela opcao palette (M8).
Public Sub Rx_OnStyle(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    Dim a As Variant
    On Error Resume Next
    If CoreOrWarn() Then
        a = StyleNames()
        If mStyleShown > 0 And UBound(a) + 1 <> mStyleShown Then
            RxMsg RX_MSG_STYLE_LIST_CHANGED, vbInformation
        ElseIf selectedIndex >= 0 And selectedIndex <= UBound(a) Then
            If mStyleLegacy Then
                RunCoreFeature "Core_SetOption", Array("palette", Rx_LegacyPaletteFor(CLng(selectedIndex)))
            Else
                RunCoreFeature "Core_StyleApply", Array(CStr(a(selectedIndex)), "chart")
            End If
        End If
    End If
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_OnStyles(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_StyleManage"
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_OnAnnotation(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_AddAnnotation", control.Tag
End Sub

' Delta clicando nas barras (Parte B): arma o modo de dois cliques no core; core anterior: aviso de atualizar.
Public Sub Rx_OnDelta(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_DeltaPick"
End Sub

Public Sub Rx_OnDeltaFormat(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_DeltaCycleFormat"
End Sub

Public Sub Rx_OnRemoveAnnotations(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_RemoveAnnotations"
End Sub

Public Sub Rx_OnRefresh(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_Refresh"
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_OnChangeSource(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_ChangeSource"
    Loader_InvalidateRibbon                 ' novos dados podem mudar o numero de series (legenda da cascata)
End Sub

' Mesma escala e Escala automatica valem sem grafico selecionado (o core avisa o que fez): os botoes ficam sempre ativos.
Public Sub Rx_OnSameScale(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_SameScale"
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_OnResetScale(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_ResetScale"
    Loader_InvalidateRibbon
End Sub

' Corte de eixo (Fase 3): o core pergunta o valor (InputBox) e aplica; ligar ou remover o corte muda o que o grafico aceita (Mesma escala,
' Remover corte de eixo), entao a faixa e invalidada.
Public Sub Rx_OnAxisBreak(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_SetAxisBreak"
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_OnAxisBreakOff(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_RemoveAxisBreak"
    Loader_InvalidateRibbon
End Sub

' Arredondar intervalo: age sobre as celulas selecionadas (sem grafico); o botao fica sempre ativo e nao muda o estado da faixa. Core anterior
' (sem o procedimento): aviso de atualizar (RunCoreFeature).
Public Sub Rx_OnRoundRange(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_RoundRange"
End Sub

' Copiar para o PowerPoint (imagem ou grafico vinculado): nao muda o grafico nem o estado da faixa. Core anterior: aviso de atualizar.
Public Sub Rx_OnPptImage(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_CopyToPowerPoint", "image"
End Sub

Public Sub Rx_OnPptChart(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_CopyToPowerPoint", "chart"
End Sub

' Titulo ligado a uma celula (1.9.0): o core pergunta a celula (InputBox) e liga o titulo do grafico ativo; core anterior: aviso de atualizar.
Public Sub Rx_OnLinkTitle(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_LinkTitle"
    Loader_InvalidateRibbon
End Sub

' Atualizar graficos no PowerPoint (Parte A): nao exige grafico selecionado; core anterior: aviso de atualizar.
Public Sub Rx_OnPptSync(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_PptSyncOpen"
End Sub

' Opcoes da atualizacao automatica (1.6.0): "autoon" e "autonew" (tag do controle). Core anterior (sem Core_PptGetFlag) ou sem core: a caixa
' fica desmarcada; o clique avisa "Atualize". Nunca propaga erro.
Public Sub Rx_GetPptFlag(control As IRibbonControl, ByRef pressed)
    On Error Resume Next
    pressed = False
    pressed = Rx_PptFlagValue(CStr(control.Tag))
End Sub

Public Sub Rx_OnPptFlag(control As IRibbonControl, pressed As Boolean)
    On Error Resume Next
    Rx_OnPptFlagCore CStr(control.Tag), pressed
End Sub

' Partes testaveis dos dois callbacks (IRibbonControl nao e criavel nos testes). Rx_PptFlagValue le o estado real do core (False sem core ou
' com um core sem o procedimento: Loader_RunCore devolve Empty); Rx_OnPptFlagCore grava e redesenha a faixa.
Public Function Rx_PptFlagValue(tag As String) As Boolean
    Dim v As Variant, p As String
    On Error Resume Next
    If Loader_CoreWorkbook() Is Nothing Then Exit Function
    p = "Core_PptGetFlag"
    If BUILD_CONFIG = "Test" Then p = p & mTestProcSuffix
    v = Loader_RunCore(p, tag)
    If Not IsEmpty(v) Then Rx_PptFlagValue = CBool(v)
    If Err.Number <> 0 Then Rx_PptFlagValue = False
End Function

Public Sub Rx_OnPptFlagCore(tag As String, pressed As Boolean)
    On Error Resume Next
    If CoreOrWarn() Then RunCoreFeature "Core_PptSetFlag", Array(tag, pressed)
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_OnSamples(control As IRibbonControl)
    On Error GoTo Fail
    If Not CoreOrWarn() Then Exit Sub
    Application.Run "'" & Loader_CoreFileName() & "'!Core_InsertSamples"
    Exit Sub
Fail:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    Loader_Log "Rx_OnSamples erro " & en & ": " & ed
    If en = 1004 And InStr(1, ed, "Core_InsertSamples", vbTextCompare) > 0 Then
        RxMsg "Os exemplos ainda não estão disponíveis nesta versão. Use 'Verificar atualizações'.", vbInformation
    Else
        RxMsg "Não foi possível inserir os exemplos: " & ed, vbExclamation
    End If
End Sub

' tutorial.html da versão instalada; se não existir, a URL pública.
Public Function Loader_TutorialPath() As String
    Dim ver As String, p As String
    ver = Loader_FindNewestCoreVersion(Loader_VersionsDir())
    If Len(ver) > 0 Then
        p = Loader_VersionsDir() & "\" & ver & "\tutorial.html"
        If Len(Dir(p)) > 0 Then Loader_TutorialPath = p: Exit Function
    End If
    Loader_TutorialPath = TUTORIAL_URL
End Function

Public Sub Rx_OnTutorial(control As IRibbonControl)
    On Error GoTo Fail
    ThisWorkbook.FollowHyperlink Loader_TutorialPath()
    Exit Sub
Fail:
    RxMsg "Não foi possível abrir o tutorial: " & Err.Description, vbExclamation
End Sub

Public Sub Rx_OnCheckUpdates(control As IRibbonControl)
    On Error Resume Next
    Updater_CheckInteractive
End Sub

Public Sub Rx_OnSettings(control As IRibbonControl)
    Dim cur As String, nv As Variant
    On Error GoTo Fail
    cur = Settings_UpdateSource()
    If BUILD_CONFIG = "Test" Then Exit Sub                ' InputBox bloquearia os testes
    nv = Application.InputBox("Fonte de atualização (URL do manifest.json ou caminho de pasta/arquivo):", APP_NAME, cur, Type:=2)
    If VarType(nv) = vbBoolean Then Exit Sub            ' cancelado
    If Len(Trim$(CStr(nv))) = 0 Then
        Settings_Delete KEY_UPDATE_SOURCE
    Else
        Settings_Set KEY_UPDATE_SOURCE, Trim$(CStr(nv))
    End If
    If Updater_PpamInstalled() Then                      ' so para quem tem o suplemento do PowerPoint instalado
        If Loader_Msg("Atualizar também o suplemento do PowerPoint automaticamente?", vbYesNo + vbQuestion) = vbYes Then
            Settings_Set KEY_PPAM_AUTO, "1"
        Else
            Settings_Set KEY_PPAM_AUTO, "0"
        End If
    End If
    RxMsg "Fonte de atualização: " & Settings_UpdateSource(), vbInformation
    Exit Sub
Fail:
    RxMsg "Não foi possível alterar as configurações: " & Err.Description, vbExclamation
End Sub

Public Sub Rx_OnAbout(control As IRibbonControl)
    On Error GoTo Fail
    RxMsg "ThinQQ" & vbLf & "Loader " & APP_VERSION & "  |  Core " & Loader_CurrentCoreVersion() & vbLf & vbLf & _
           "Gráficos de apresentação para Excel." & vbLf & "Pasta: " & Loader_Root(), vbInformation
    Exit Sub
Fail:
    RxMsg "Não foi possível exibir as informações: " & Err.Description, vbExclamation
End Sub
