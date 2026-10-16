Attribute VB_Name = "modPptTqq"
Option Explicit
' Suplemento do PowerPoint do ThinQQ (ThinQQ.ppam; spec 2026-10-14 secao 4.3). Ao abrir uma apresentacao com graficos ThinQQ vinculados
' (etiqueta TQQ_CHART, gravada pelo core ao colar) pergunta (padrao), atualiza sempre ou nao faz nada (opcao "Ao abrir") e pede ao core do
' ThinQQ no Excel para atualizar dados e anotacoes (Core_PptSyncPresentation). Excel: o aberto, se ja tem o core; senao uma instancia NOVA e
' invisivel (o Excel e de varias instancias: nunca a do usuario), com o core mais novo de %LOCALAPPDATA%\ThinQQ\versions, fechada no fim.
' Sem o core: so os dados (vinculo nativo), nunca com a pasta inexistente (o Update travaria o PowerPoint). Ligacao tardia; sem Declare.
Public Const PPT_APP_NAME As String = "ThinQQ"
Public Const PPT_SECTION As String = "PowerPoint"
Public Const PPT_KEY_OPENMODE As String = "OpenMode"
Public Const PPT_TAG_CHART As String = "TQQ_CHART"
Public Const PPT_TAG_BOOK As String = "TQQ_BOOK"
Public Const PPT_MSG_NONE As String = "Esta apresentação não tem gráficos do ThinQQ vinculados."
Public Const PPT_MSG_NO_PRES As String = "Abra uma apresentação primeiro."
Public Const PPT_MSG_NO_EXCEL As String = "O Excel não está instalado (ou não pôde ser iniciado): os gráficos do ThinQQ não foram atualizados."
Public Const PPT_MSG_DATA_ONLY As String = "O ThinQQ não está instalado no Excel deste computador: só os dados foram atualizados; setas e outras anotações do ThinQQ ficaram como estavam."
Public Const PPT_MSG_OLD_CORE As String = "Atualize o ThinQQ no Excel para atualizar também as setas."
Public Const PPT_MSG_BUSY As String = "O Excel está ocupado (uma célula em edição ou uma caixa de diálogo aberta). Termine o que está fazendo no Excel e tente de novo."
Public Const PPT_MAX_SOURCES_SHOWN As Long = 5                ' pastas de origem listadas na pergunta ao abrir (depois, "...")
Public Const PPT_MAX_PATH_SHOWN As Long = 150                 ' caracteres de cada pasta na pergunta (cortada no meio)
Public Const PPT_MAX_ASK_LEN As Long = 900                    ' tamanho maximo da pergunta (o MsgBox corta perto de 1024)
Public Const PPT_REQUIRED_API As Long = 1                   ' Core_PptApiVersion minima (spec 4.9); core sem ela ou menor = so os dados
Public Const PPT_TAG_AUTO As String = "TQQ_AUTO"
Public Const PPT_TAG_INDEX As String = "TQQ_AUTOIDX"
Public Const PPT_API_SHAPES As Long = 2                     ' Core_PptSyncShapes (core 1.6.0); conferida POR RECURSO (ao abrir basta a 1)
Public Const PPT_IDX_MAX_ENTRIES As Long = 2000             ' = PPTIDX_MAX_ENTRIES do core
Public Const PPT_IDX_MAX_LEN As Long = 1000000              ' = PPTIDX_MAX_LEN do core
Private Const PPT_IDX_FS As String = "|"                    ' indice: separador de campos (= PPTIDX_FS do core)
Private Const PPT_IDX_ES As String = vbLf                   ' indice: separador de entradas (= PPTIDX_ES do core)
Public Const PPT_MSG_SELECT As String = "Selecione um ou mais gráficos do ThinQQ vinculados ao Excel."
Public Const PPT_MSG_AUTO_OLDCORE As String = "Atualize o ThinQQ no Excel para a atualização automática funcionar (a marca foi gravada)."
Public Const PPT_MSG_AUTO_NOCORE As String = "Marcado para atualizar automaticamente. Para os gráficos acompanharem a pasta de origem, abra o Excel com o ThinQQ (1.6.0 ou mais novo)."
Public Const PPT_MSG_AUTO_OFF As String = "A atualização automática está desligada no Excel: aba ThinQQ > PowerPoint > Atualização automática ligada."
Public Const PPT_MSG_SEL_OLDCORE As String ="Atualize o ThinQQ no Excel para atualizar só os gráficos selecionados (desta vez a apresentação inteira foi atualizada)."
Public Const PPT_TAG_DATA As String = "TQQ_DATA"
Public Const PPT_TAG_DATABOOK As String = "TQQ_DATABOOK"
Public Const PPT_DATA_SUFFIX As String = " - dados ThinQQ.xlsx"
Public Const PPT_API_DATA As Long = 3                       ' Core_PptInsertChart, Core_PptEditData, Core_PptDataSave (core 1.8.0)
Public Const PPT_MSG_SAVE_FIRST As String = "Salve a apresentação primeiro: os dados do gráfico ficam numa pasta do Excel ao lado dela (""<nome> - dados ThinQQ.xlsx"")."
Public Const PPT_MSG_WEB_PRES As String = "A apresentação está aberta por um endereço da Web (OneDrive/SharePoint). Salve-a numa pasta do computador (ou na pasta sincronizada do OneDrive) para inserir gráficos do ThinQQ."
Public Const PPT_MSG_DATA_NOCORE As String = "O ThinQQ não está instalado no Excel deste computador: não é possível criar nem editar gráficos do ThinQQ."
Public Const PPT_MSG_OLD_CORE_DATA As String = "Atualize o ThinQQ no Excel para criar gráficos e editar dados pelo PowerPoint."
Public Const PPT_MSG_NO_EXCEL_DATA As String = "O Excel não está instalado (ou não pôde ser iniciado): o ThinQQ precisa dele para criar gráficos e editar dados."
Public Const PPT_API_CMD As Long = 4                        ' Core_PptCommand, Core_PptCategories (core 1.9.0)
Public Const PPT_TAG_OPTS As String = "TQQ_OPTS"            ' estado das opcoes gravado pelo core na forma ("labels=1;...;caps=labels,...,")
Public Const PPT_MSG_OLD_CORE_CMD As String = "Atualize o ThinQQ no Excel para mudar o gráfico pelo PowerPoint."
Public Const PPT_MSG_CATS As String = "Escolha duas categorias diferentes do gráfico, a primeira antes da segunda."
Public Const PPT_MSG_PACK_SAVE_FIRST As String = "Salve a apresentação numa pasta do computador antes de empacotar."
Public Const PPT_PACK_SUFFIX As String = " (ThinQQ)"            ' pasta do pacote: "<pasta da apresentacao>\<nome> (ThinQQ)"
Public Const PPT_API_TEXT As Long = 5                       ' Core_PptLinkText (core 1.9.0)
Public Const PPT_TAG_TEXT As String = "TQQ_TEXT"            ' forma de texto vinculado a celula ("1"; = PPTTEXT_TAG do core)
Public Const PPT_MSG_OPEN_EXCEL As String = "Abra a pasta no Excel (com o ThinQQ), selecione a célula com o valor e clique de novo em Texto vinculado ao Excel."
Public Const PPTTEXT_MSG_NOTEXT_PPT As String = "Selecione uma caixa de texto (ou parte do texto dela)."
Public Const PPT_MSG_OLD_CORE_TEXT As String = "Atualize o ThinQQ no Excel para vincular texto pelo PowerPoint."
Public Const PPT_MSG_TEXT_GROUP As String = "O texto está dentro de um grupo: desagrupe as formas (ou use uma caixa de texto fora do grupo) para vincular."
Public Const PPT_MSG_TEXT_OK As String = "Texto vinculado à célula do Excel."
Public Const PPT_MSG_TEXT_SAVE As String = "Salve a pasta de trabalho no Excel: assim o vínculo acompanha linhas e colunas inseridas antes da célula."
Private Const RPC_E_CALL_REJECTED As Long = -2147418111
Private Const RPC_E_RETRYLATER As Long = -2147417846
Private mEvents As clsPptEvents
Private mSeen As Collection
Private mTestXl As Object
Private mTestCoreDir As String
Private mTestCreated As Boolean                 ' gancho de teste: o Excel injetado conta como criado por nos (Quit no fim)
Private mTestCoreInfo As String                 ' gancho de teste: "caminho|ro" do core aberto pelo suplemento
Private mTestOnOpenCalls As Long                ' gancho de teste: chamadas de PptTqq_OnOpen (eventos e Auto_Open)
Private mLastMsg As String
Private mLastLog As String
Private mBusy As Boolean                        ' reentrada: pergunta/sincronizacao em andamento (eventos chegam durante o MsgBox)
Private mEditXl As Object                       ' Excel criado aqui para criar/editar dados (fica com o usuario; atalho para reencontra-lo)
Private mHintShown As Boolean                   ' aviso de primeira vez de "Inserir grafico" ja mostrado nesta sessao
Private mTestTrail As String                    ' gancho de teste: ordem dos passos de inserir ("hint;edit;")
Private mLastHint As String                     ' gancho de teste: texto do aviso de primeira vez
Private mTestProbes As Long                     ' gancho de teste: FileExists do caminho antigo em PptTqq_RelinkData
Private mTestPickA As Long                      ' gancho de teste: respostas de PptTqq_PickTwo (de, ate)
Private mTestPickB As Long
Private mTestPackProbes As Long                 ' gancho de teste: FileExists de pastas de dados em PptTqq_Package
Private mTestInvalidates As Long                ' gancho de teste: chamadas de PptRx_InvalidateElem

' Carga do suplemento: liga os eventos do PowerPoint e trata as apresentacoes ja abertas (o projeto de teste .pptm nao liga).
Public Sub Auto_Open()
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then Exit Sub
    Set mEvents = New clsPptEvents
    Set mEvents.App = Application
    PptTqq_ScanOpen
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Set mEvents = Nothing
End Sub

' Apresentacoes ja abertas quando o suplemento carrega: as mesmas regras da abertura (inclusive a guarda de janela).
Public Sub PptTqq_ScanOpen()
    Dim i As Long, n As Long
    On Error Resume Next
    n = Application.Presentations.Count
    For i = 1 To n
        PptTqq_OnOpen Application.Presentations(i)
    Next
End Sub

' Gancho de teste: Excel injetado, pasta das versoes do core e se o Excel injetado conta como criado pelo suplemento (asCreated: abre o
' core da pasta e da Quit no fim). Inerte fora do build Test.
Public Sub PptTqq_SetExcelForTest(xl As Object, coreDir As String, Optional asCreated As Boolean = False)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    Set mTestXl = xl
    mTestCoreDir = coreDir
    mTestCreated = asCreated
    mTestCoreInfo = ""
End Sub

' Gancho de teste: zera todos os ganchos e o estado da sessao (vistas, reentrada, contadores). Inerte fora do build Test.
Public Sub PptTqq_ResetTestHooks()
    If BUILD_CONFIG <> "Test" Then Exit Sub
    Set mTestXl = Nothing
    mTestCoreDir = ""
    mTestCreated = False
    mTestCoreInfo = ""
    mTestOnOpenCalls = 0
    mBusy = False
    Set mSeen = Nothing
    mLastLog = ""
    mHintShown = False
    mTestTrail = ""
    mLastHint = ""
    mTestProbes = 0
    mTestPickA = 0
    mTestPickB = 0
    mTestPackProbes = 0
    mTestInvalidates = 0
End Sub

Public Function PptTqq_RelinkProbesForTest() As Long
    PptTqq_RelinkProbesForTest = mTestProbes
End Function

Public Function PptTqq_CoreInfoForTest() As String
    PptTqq_CoreInfoForTest = mTestCoreInfo
End Function

Public Function PptTqq_OnOpenCallsForTest() As Long
    PptTqq_OnOpenCallsForTest = mTestOnOpenCalls
End Function

' Log silencioso (caminho dos eventos): %LOCALAPPDATA%\ThinQQ\logs\thinqq.log se a pasta existir (nunca cria); no build Test so memoria.
Public Sub PptTqq_Log(msg As String)
    Dim f As Integer, d As String
    mLastLog = msg
    If BUILD_CONFIG = "Test" Then Exit Sub
    On Error Resume Next
    d = Environ$("LOCALAPPDATA") & "\ThinQQ\logs"
    If Len(Dir$(d, vbDirectory)) = 0 Then Exit Sub
    f = FreeFile
    Open d & "\thinqq.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & "ppt" & vbTab & msg
    Close #f
End Sub

Public Function PptTqq_LastLog() As String
    PptTqq_LastLog = mLastLog
End Function

Public Function PptTqq_LastMsg() As String
    PptTqq_LastMsg = mLastMsg
End Function

Public Sub PptTqq_ClearLastMsg()
    mLastMsg = ""
End Sub

' Unico ponto de MsgBox de aviso (log-only no build Test).
Public Sub PptTqq_Msg(msg As String, Optional style As VbMsgBoxStyle = vbInformation)
    mLastMsg = msg
    If BUILD_CONFIG = "Test" Then Exit Sub
    MsgBox msg, style, PPT_APP_NAME
End Sub

' Unico ponto de pergunta (no build Test registra e responde Nao).
Private Function PptTqq_Ask(msg As String) As VbMsgBoxResult
    mLastMsg = msg
    If BUILD_CONFIG = "Test" Then PptTqq_Ask = vbNo: Exit Function
    PptTqq_Ask = MsgBox(msg, vbQuestion + vbYesNo, PPT_APP_NAME)
End Function

Public Function PptTqq_NormalizeMode(m As String) As String
    Select Case LCase$(Trim$(m))
        Case "always": PptTqq_NormalizeMode = "always"
        Case "never": PptTqq_NormalizeMode = "never"
        Case Else: PptTqq_NormalizeMode = "ask"
    End Select
End Function

' Modo "Ao abrir" (HKCU\Software\VB and VBA Program Settings\ThinQQ\PowerPoint\OpenMode). Testes: sempre "ask", sem ler o registro.
Public Function PptTqq_OpenMode() As String
    If BUILD_CONFIG = "Test" Then PptTqq_OpenMode = "ask": Exit Function
    PptTqq_OpenMode = PptTqq_NormalizeMode(GetSetting(PPT_APP_NAME, PPT_SECTION, PPT_KEY_OPENMODE, "ask"))
End Function

Public Sub PptTqq_SetOpenMode(m As String)
    If BUILD_CONFIG = "Test" Then Exit Sub
    SaveSetting PPT_APP_NAME, PPT_SECTION, PPT_KEY_OPENMODE, PptTqq_NormalizeMode(m)
End Sub

' O que fazer ao abrir: "" (nada), "ask" ou "sync". hasRemoteSource: alguma pasta de origem e de rede (\\servidor\...) ou de endereco
' http(s) e ainda nao esta aberta no Excel: "Atualizar sempre" volta a perguntar (a apresentacao escolhe quais pastas a atualizacao abre;
' uma origem remota nao e aberta sem o usuario ver o caminho). Pura.
Public Function PptTqq_OnOpenAction(mode As String, isReadOnly As Boolean, nTagged As Long, alreadySeen As Boolean, hasRemoteSource As Boolean) As String
    If nTagged = 0 Or alreadySeen Or isReadOnly Or mode = "never" Then Exit Function
    If mode = "always" And Not hasRemoteSource Then PptTqq_OnOpenAction = "sync" Else PptTqq_OnOpenAction = "ask"
End Function

' Pergunta ao abrir. sources (opcional): as pastas de origem distintas, listadas antes da dica (PptTqq_SourcesText: remotas primeiro, no
' maximo PPT_MAX_SOURCES_SHOWN, cada uma limpa e com no maximo PPT_MAX_PATH_SHOWN caracteres). A pergunta inteira fica em
' PPT_MAX_ASK_LEN caracteres (o MsgBox corta perto de 1024): a frase da pergunta e a dica sempre aparecem.
Public Function PptTqq_AskText(n As Long, Optional sources As Collection) As String
    Dim head As String, tail As String
    head = "Esta apresentação tem " & n & " gráfico(s) do ThinQQ vinculado(s) ao Excel. Atualizar agora com os dados atuais?"
    tail = vbLf & vbLf & "Para não perguntar mais: aba ThinQQ > Ao abrir."
    PptTqq_AskText = head & PptTqq_SourcesText(sources, PPT_MAX_ASK_LEN - Len(head) - Len(tail)) & tail
End Function

' Bloco "Pastas de origem:" da pergunta (vazio sem pastas): remotas primeiro (na ordem em que aparecem), depois as locais; para no
' PPT_MAX_SOURCES_SHOWN-esimo caminho ou quando o proximo passaria do limite budget, com "..." se sobrou alguma. Pura.
Private Function PptTqq_SourcesText(sources As Collection, budget As Long) As String
    Dim ordered As New Collection, v As Variant, s As String, ln As String, k As Long, pass As Long
    If sources Is Nothing Then Exit Function
    If sources.Count = 0 Then Exit Function
    For pass = 1 To 2
        For Each v In sources
            If PptTqq_IsRemotePath(CStr(v)) = (pass = 1) Then ordered.Add PptTqq_CleanPath(CStr(v))
        Next
    Next
    s = vbLf & vbLf & "Pastas de origem:"
    For Each v In ordered
        ln = vbLf & CStr(v)
        If k >= PPT_MAX_SOURCES_SHOWN Or Len(s) + Len(ln) + 4 > budget Then s = s & vbLf & "...": Exit For
        s = s & ln
        k = k + 1
    Next
    PptTqq_SourcesText = s
End Function

' Caminho para exibir: cada caractere de controle (codigo < 32 ou 127) vira espaco, espacos repetidos viram um, e acima de
' PPT_MAX_PATH_SHOWN caracteres fica o comeco e o fim com "..." no meio. Pura.
Public Function PptTqq_CleanPath(p As String) As String
    Dim i As Long, c As Long, s As String, keepHead As Long
    s = p
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If (c >= 0 And c < 32) Or c = 127 Then Mid$(s, i, 1) = " "
    Next
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)
    If Len(s) > PPT_MAX_PATH_SHOWN Then
        keepHead = (PPT_MAX_PATH_SHOWN - 3) \ 2
        s = Left$(s, keepHead) & "..." & Right$(s, PPT_MAX_PATH_SHOWN - 3 - keepHead)
    End If
    PptTqq_CleanPath = s
End Function

' Origem remota: caminho de rede (\\servidor\..., inclusive \\?\UNC\ e WebDAV \\servidor@SSL\DavWWWRoot\), com barras normais ou
' mistas (//servidor/...), endereco http(s) ou URL file: (sempre remota, por cautela). \\?\C:\ (unidade local com prefixo longo) e
' letras de unidade sao locais: uma unidade de rede mapeada (Z:\) nao e distinguivel aqui (limite documentado). Pura.
Public Function PptTqq_IsRemotePath(p As String) As Boolean
    Dim t As String
    t = Replace(LCase$(Trim$(p)), "/", "\")
    If Left$(t, 5) = "file:" Or Left$(t, 7) = "http:\\" Or Left$(t, 8) = "https:\\" Then PptTqq_IsRemotePath = True: Exit Function
    If Left$(t, 2) <> "\\" Then Exit Function
    If Left$(t, 8) = "\\?\unc\" Then PptTqq_IsRemotePath = True: Exit Function
    If (Left$(t, 4) = "\\?\" Or Left$(t, 4) = "\\.\") And Mid$(t, 6, 1) = ":" Then Exit Function    ' \\?\C:\ ou \\.\C:\: unidade local
    PptTqq_IsRemotePath = True
End Function

' Alguma pasta de origem remota que nao esta entre as abertas no Excel (names: FullName das pastas abertas). Pura.
Public Function PptTqq_HasRemoteSource(sources As Collection, names As Collection) As Boolean
    Dim v As Variant
    If sources Is Nothing Then Exit Function
    For Each v In sources
        If PptTqq_IsRemotePath(CStr(v)) Then
            If Not PptTqq_InNames(CStr(v), names) Then PptTqq_HasRemoteSource = True: Exit Function
        End If
    Next
End Function

Private Function PptTqq_InNames(p As String, names As Collection) As Boolean
    Dim n As Variant
    If names Is Nothing Then Exit Function
    For Each n In names
        If StrComp(CStr(n), Trim$(p), vbTextCompare) = 0 Then PptTqq_InNames = True: Exit Function
    Next
End Function

' Mesma regra do core (PptSync_ReachableAmong): aberta no Excel (FullName igual; vale para o endereco https) ou, se nao for http(s),
' arquivo existente no disco. Nunca propaga erro.
Public Function PptTqq_ReachableAmong(path As String, names As Collection) As Boolean
    Dim p As String
    On Error GoTo Done
    p = Trim$(path)
    If Len(p) = 0 Then Exit Function
    If PptTqq_InNames(p, names) Then PptTqq_ReachableAmong = True: Exit Function
    If LCase$(Left$(p, 7)) = "http://" Or LCase$(Left$(p, 8)) = "https://" Then Exit Function
    PptTqq_ReachableAmong = PptTqq_FileExists(p)
Done:
End Function

' FullName das pastas abertas num Excel (vazio se xl e Nothing ou nao responde). Nunca propaga erro.
Private Function PptTqq_WorkbookNames(xl As Object) As Collection
    Dim out As New Collection, wb As Object
    Set PptTqq_WorkbookNames = out
    If xl Is Nothing Then Exit Function
    On Error Resume Next
    For Each wb In xl.Workbooks
        out.Add CStr(wb.FullName)
    Next
    Err.Clear
End Function

' Pastas abertas no Excel em execucao, sem nunca criar um (build Test: so o Excel injetado). Ocupado ou ausente: vazio.
Private Function PptTqq_ExcelOpenNames() As Collection
    Dim xl As Object
    If BUILD_CONFIG = "Test" Then Set PptTqq_ExcelOpenNames = PptTqq_WorkbookNames(mTestXl): Exit Function
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    Err.Clear
    On Error GoTo 0
    Set PptTqq_ExcelOpenNames = PptTqq_WorkbookNames(xl)
End Function

Private Function PptTqq_Tag(shp As Object, key As String) As String
    On Error Resume Next
    PptTqq_Tag = CStr(shp.Tags.Item(key))
    Err.Clear
End Function

' Formas soltas marcadas pelo ThinQQ na apresentacao: graficos (TQQ_CHART) e, desde a 1.9.0, textos vinculados (TQQ_TEXT = "1": uma
' apresentacao so com textos vinculados tambem pergunta/atualiza ao abrir e pelo botao). sources (opcional): recebe as pastas de origem
' distintas (sem diferenciar maiusculas), na ordem em que aparecem: PptTqq_SourcePath(LinkFormat.SourceFullName, TQQ_BOOK); so le o
' vinculo e as etiquetas (nunca o arquivo).
Public Function PptTqq_CountTagged(pres As Object, Optional sources As Collection) As Long
    Dim sl As Object, shp As Object, p As String
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If Len(PptTqq_Tag(shp, PPT_TAG_CHART)) > 0 Or PptTqq_Tag(shp, PPT_TAG_TEXT) = "1" Then
                PptTqq_CountTagged = PptTqq_CountTagged + 1
                If Not sources Is Nothing Then
                    p = PptTqq_SourcePath(PptTqq_LinkSource(shp), PptTqq_Tag(shp, PPT_TAG_BOOK))
                    If Len(p) > 0 Then
                        On Error Resume Next
                        sources.Add p, LCase$(p)                ' repetida: chave ja existe, ignorada
                        Err.Clear
                        On Error GoTo 0
                    End If
                End If
            End If
        Next
    Next
End Function

' LinkFormat.SourceFullName da forma ("" se nao ha vinculo legivel). Nunca propaga erro.
Private Function PptTqq_LinkSource(shp As Object) As String
    On Error Resume Next
    PptTqq_LinkSource = CStr(shp.LinkFormat.SourceFullName)
    Err.Clear
End Function

' Mesma regra do core (PptSync_SourcePath). Pura.
Public Function PptTqq_SourcePath(linkFull As String, tagBook As String) As String
    Dim s As String, p As Long
    s = Trim$(linkFull)
    p = InStr(s, "!")
    If p > 0 Then s = Left$(s, p - 1)
    If Len(s) = 0 Then s = Trim$(tagBook)
    PptTqq_SourcePath = s
End Function

Private Function PptTqq_IsVersion(s As String) As Boolean
    Dim p() As String, i As Long
    p = Split(s, ".")
    If UBound(p) <> 2 Then Exit Function
    For i = 0 To 2
        If Len(p(i)) = 0 Or Not IsNumeric(p(i)) Or InStr(p(i), ",") > 0 Or InStr(p(i), "-") > 0 Then Exit Function
    Next
    PptTqq_IsVersion = True
End Function

' Compara x.y.z numericamente (1 se a > b, -1 se a < b, 0 iguais).
Private Function PptTqq_CompareVersion(a As String, b As String) As Long
    Dim pa() As String, pb() As String, i As Long
    pa = Split(a, "."): pb = Split(b, ".")
    For i = 0 To 2
        If CLng(pa(i)) > CLng(pb(i)) Then PptTqq_CompareVersion = 1: Exit Function
        If CLng(pa(i)) < CLng(pb(i)) Then PptTqq_CompareVersion = -1: Exit Function
    Next
End Function

' Maior versao x.y.z numa lista "a|b|c" (nomes de pastas); "" sem nenhuma valida. Pura.
Public Function PptTqq_NewestVersion(names As String) As String
    Dim p() As String, i As Long, best As String
    p = Split(names, "|")
    For i = LBound(p) To UBound(p)
        If PptTqq_IsVersion(p(i)) Then
            If Len(best) = 0 Then
                best = p(i)
            ElseIf PptTqq_CompareVersion(p(i), best) > 0 Then
                best = p(i)
            End If
        End If
    Next
    PptTqq_NewestVersion = best
End Function

Private Function PptTqq_CoreFileName() As String
    PptTqq_CoreFileName = "ThinQQ.core" & BUILD_SUFFIX & ".xlam"
End Function

Private Function PptTqq_CoreDir() As String
    If BUILD_CONFIG = "Test" Then PptTqq_CoreDir = mTestCoreDir: Exit Function
    PptTqq_CoreDir = Environ$("LOCALAPPDATA") & "\ThinQQ\versions"
End Function

' Pura: "nome:1|nome:0|..." (1 = a pasta tem o arquivo do core) -> versoes x.y.z COM o arquivo, da mais nova para a mais velha ("a|b").
' Mesma regra do loader (Loader_ListCoreVersions): pasta parcial, sem o arquivo, fica de fora.
Public Function PptTqq_CoreCandidates(listing As String) As String
    Dim p() As String, a() As String, i As Long, j As Long, n As Long, k As Long, nm As String, t As String, s As String
    p = Split(listing, "|")
    ReDim a(0 To UBound(p) + 1)
    For i = LBound(p) To UBound(p)
        k = InStrRev(p(i), ":")
        If k > 1 Then
            nm = Left$(p(i), k - 1)
            If Mid$(p(i), k + 1) = "1" And PptTqq_IsVersion(nm) Then
                a(n) = nm
                n = n + 1
            End If
        End If
    Next
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If PptTqq_CompareVersion(a(j), a(i)) > 0 Then t = a(i): a(i) = a(j): a(j) = t
        Next
    Next
    For i = 0 To n - 1
        If i > 0 Then s = s & "|"
        s = s & a(i)
    Next
    PptTqq_CoreCandidates = s
End Function

' Versoes candidatas na pasta d (subpastas com o arquivo do core), da mais nova para a mais velha.
Public Function PptTqq_ListCoreCandidates(d As String) As String
    Dim fso As Object, sf As Object, s As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function
    If Not fso.FolderExists(d) Then Exit Function
    For Each sf In fso.GetFolder(d).SubFolders
        s = s & sf.Name & ":" & IIf(fso.FileExists(sf.Path & "\" & PptTqq_CoreFileName()), "1", "0") & "|"
    Next
    Err.Clear
    PptTqq_ListCoreCandidates = PptTqq_CoreCandidates(s)
End Function

' Pura: resultado do Core_Ping. 1 = core respondeu; -1 = Excel ocupado (RPC); 0 = sem core.
Public Function PptTqq_PingClass(en As Long, v As String) As Long
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Then PptTqq_PingClass = -1: Exit Function
    If en = 0 And v = "pong" Then PptTqq_PingClass = 1
End Function

Private Function PptTqq_PingState(xl As Object) As Long
    Dim v As Variant, en As Long, s As String
    On Error Resume Next
    v = xl.Run("'" & PptTqq_CoreFileName() & "'!Core_Ping")
    en = Err.Number
    Err.Clear
    s = CStr(v)
    Err.Clear
    PptTqq_PingState = PptTqq_PingClass(en, s)
End Function

Private Sub PptTqq_RaiseBusy()
    Err.Raise RPC_E_CALL_REJECTED, "PptTqq", "Excel ocupado"
End Sub

' Excel a usar: no build Test, so o injetado (created conforme o gancho). Senao: o aberto (GetObject) SE ele tem o core (ocupado: erro de
' ocupado, sem cair para outra instancia); senao uma instancia nova e invisivel (created). O retorno e atribuido antes de configurar: a
' instancia criada nunca fica sem dono. noCreate: so o Excel ja aberto (com ou sem o core), nunca uma instancia nova (Nothing se nao ha).
Private Function PptTqq_GetExcel(ByRef created As Boolean, Optional noCreate As Boolean = False) As Object
    Dim xl As Object, st As Long
    created = False
    If BUILD_CONFIG = "Test" Then
        created = mTestCreated And Not noCreate
        Set PptTqq_GetExcel = mTestXl
        Exit Function
    End If
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    Err.Clear
    On Error GoTo 0
    If Not xl Is Nothing Then
        st = PptTqq_PingState(xl)
        If st = 1 Then Set PptTqq_GetExcel = xl: Exit Function
        If st = -1 Then PptTqq_RaiseBusy
        If noCreate Then Set PptTqq_GetExcel = xl: Exit Function
    End If
    If noCreate Then Exit Function
    Set xl = Nothing
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    Err.Clear
    If xl Is Nothing Then Exit Function
    created = True
    Set PptTqq_GetExcel = xl
    xl.Visible = False
    xl.DisplayAlerts = False
    Err.Clear
End Function

' Abre um core (somente leitura) e confere o Core_Ping; se nao responder, fecha o que abriu.
Private Function PptTqq_TryOpenCore(xl As Object, p As String) As Boolean
    Dim wb As Object
    On Error Resume Next
    Set wb = xl.Workbooks.Open(p, 0, True)
    Err.Clear
    If PptTqq_PingState(xl) = 1 Then
        PptTqq_TryOpenCore = True
        If BUILD_CONFIG = "Test" Then mTestCoreInfo = CStr(xl.Workbooks(PptTqq_CoreFileName()).FullName) & "|" & IIf(xl.Workbooks(PptTqq_CoreFileName()).ReadOnly, "ro", "rw")
        Err.Clear
        Exit Function
    End If
    If wb Is Nothing Then Set wb = xl.Workbooks(PptTqq_CoreFileName())
    If Not wb Is Nothing Then wb.Close False
    Err.Clear
End Function

' Nome do core respondendo no Excel ("" sem core). Na instancia criada por nos, abre o core instalado: da versao mais nova para a mais velha,
' a primeira que responder (como o loader: versao nova que nao carrega cai para a anterior).
Private Function PptTqq_EnsureCore(xl As Object, created As Boolean) As String
    Dim st As Long, d As String, cands As String, c() As String, i As Long
    st = PptTqq_PingState(xl)
    If st = -1 Then PptTqq_RaiseBusy
    If st = 1 Then PptTqq_EnsureCore = PptTqq_CoreFileName(): Exit Function
    If Not created Then Exit Function
    d = PptTqq_CoreDir()
    If Len(d) = 0 Then Exit Function
    cands = PptTqq_ListCoreCandidates(d)
    If Len(cands) = 0 Then Exit Function
    c = Split(cands, "|")
    For i = 0 To UBound(c)
        If PptTqq_TryOpenCore(xl, d & "\" & c(i) & "\" & PptTqq_CoreFileName()) Then
            PptTqq_EnsureCore = PptTqq_CoreFileName()
            Exit Function
        End If
        PptTqq_Log "PptTqq_EnsureCore: core " & c(i) & " nao carregou; tentando a versao anterior"
    Next
End Function

' Quit so na instancia que o suplemento criou.
Private Sub PptTqq_Release(xl As Object, created As Boolean)
    On Error Resume Next
    If created And Not xl Is Nothing Then
        xl.DisplayAlerts = False
        xl.Quit
        If BUILD_CONFIG = "Test" Then Set mTestXl = Nothing
    End If
    Set xl = Nothing
End Sub

Public Function PptTqq_ErrorText(en As Long, ed As String) As String
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Then
        PptTqq_ErrorText = PPT_MSG_BUSY
    ElseIf en = 429 Then
        PptTqq_ErrorText = PPT_MSG_NO_EXCEL
    Else
        PptTqq_ErrorText = "Não foi possível atualizar os gráficos do ThinQQ: " & ed & " (" & en & ")"
    End If
End Function

' Modo degradado (sem o core): so os dados pelo vinculo nativo, e so dos graficos cuja pasta e alcancavel pela regra do core
' (PptTqq_ReachableAmong: endereco https so se estiver aberto no Excel; names = pastas abertas no Excel usado).
Public Function PptTqq_DataOnly(pres As Object, Optional names As Collection) As String
    Dim sl As Object, shp As Object, n As Long, nMiss As Long, p As String, lf As String
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If Len(PptTqq_Tag(shp, PPT_TAG_CHART)) > 0 Then
                lf = ""
                On Error Resume Next
                lf = CStr(shp.LinkFormat.SourceFullName)
                Err.Clear
                On Error GoTo 0
                p = PptTqq_SourcePath(lf, PptTqq_Tag(shp, PPT_TAG_BOOK))
                If PptTqq_ReachableAmong(p, names) Then
                    On Error Resume Next
                    shp.LinkFormat.Update
                    If Err.Number = 0 Then n = n + 1
                    Err.Clear
                    On Error GoTo 0
                Else
                    nMiss = nMiss + 1
                End If
            End If
        Next
    Next
    PptTqq_DataOnly = "Dados atualizados em " & n & " gráfico(s)" & IIf(nMiss > 0, "; " & nMiss & " sem a pasta de origem (ficaram como estavam)", "") & "."
End Function

Private Function PptTqq_FileExists(p As String) As Boolean
    On Error Resume Next
    If Len(p) = 0 Then Exit Function
    PptTqq_FileExists = CreateObject("Scripting.FileSystemObject").FileExists(p)
    Err.Clear
End Function

' Versao da API do core (0 = core sem Core_PptApiVersion, anterior a 1.4.0). Excel ocupado: erro de ocupado.
Private Function PptTqq_CoreApi(xl As Object, core As String) As Long
    Dim v As Variant, en As Long
    On Error Resume Next
    v = xl.Run("'" & core & "'!Core_PptApiVersion")
    en = Err.Number
    Err.Clear
    On Error GoTo 0
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Then PptTqq_RaiseBusy
    If en = 0 And IsNumeric(v) Then PptTqq_CoreApi = CLng(v)
End Function

' Contrato (spec 4.9): so chama a API se o core a tem na versao exigida. Pura.
Public Function PptTqq_ApiCompatible(coreApi As Long) As Boolean
    PptTqq_ApiCompatible = (coreApi >= PPT_REQUIRED_API)
End Function

' Core_PptSyncPresentation no Excel. Excel ocupado: o erro sobe (aviso proprio). Core antigo ou API menor: so os dados + aviso. Core com a
' API que falha na chamada: o texto do erro (nao o de core antigo). api: a versao ja lida (-1 = ler agora).
Private Function PptTqq_RunCoreSync(xl As Object, core As String, pres As Object, force As Boolean, Optional api As Long = -1) As String
    Dim v As Variant, en As Long, ed As String
    If api < 0 Then api = PptTqq_CoreApi(xl, core)
    If Not PptTqq_ApiCompatible(api) Then PptTqq_RunCoreSync = PptTqq_DataOnly(pres, PptTqq_WorkbookNames(xl)) & vbLf & PPT_MSG_OLD_CORE: Exit Function
    On Error Resume Next
    v = xl.Run("'" & core & "'!Core_PptSyncPresentation", pres, force)
    en = Err.Number: ed = Err.Description
    Err.Clear
    On Error GoTo 0
    If en = 0 Then PptTqq_RunCoreSync = CStr(v): Exit Function
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Then Err.Raise en, "PptTqq", ed
    PptTqq_RunCoreSync = PptTqq_ErrorText(en, ed)
End Function

' Atualiza os graficos ThinQQ da apresentacao (force: inclusive os que nao mudaram). Aviso final com o resumo. Nunca propaga erro.
' Reentrada (evento durante outra pergunta/sincronizacao): ignora.
Public Sub PptTqq_Sync(pres As Object, force As Boolean)
    If mBusy Then Exit Sub
    mBusy = True
    On Error Resume Next
    PptTqq_DoSync pres, force
    mBusy = False
End Sub

Private Sub PptTqq_DoSync(pres As Object, force As Boolean)
    Dim xl As Object, created As Boolean, core As String, msg As String, en As Long, ed As String
    On Error GoTo Fail
    If PptTqq_CountTagged(pres) = 0 Then PptTqq_Msg PPT_MSG_NONE: Exit Sub
    Set xl = PptTqq_GetExcel(created)
    If xl Is Nothing Then PptTqq_Msg PPT_MSG_NO_EXCEL, vbExclamation: Exit Sub
    core = PptTqq_EnsureCore(xl, created)
    If Len(core) > 0 Then
        msg = PptTqq_RunCoreSync(xl, core, pres, force)
    Else
        msg = PptTqq_DataOnly(pres, PptTqq_WorkbookNames(xl)) & vbLf & PPT_MSG_DATA_ONLY
    End If
    PptTqq_Release xl, created
    PptTqq_Msg msg, vbInformation
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptTqq_Release xl, created
    PptTqq_Msg PptTqq_ErrorText(en, ed), vbExclamation
End Sub

Private Function PptTqq_Seen(key As String) As Boolean
    Dim v As Variant
    If mSeen Is Nothing Then Exit Function
    On Error Resume Next
    v = mSeen(key)
    PptTqq_Seen = (Err.Number = 0)
    Err.Clear
End Function

Private Sub PptTqq_MarkSeen(key As String)
    If mSeen Is Nothing Then Set mSeen = New Collection
    On Error Resume Next
    mSeen.Add key, key
    Err.Clear
End Sub

' Evento PresentationCloseFinal: esquece a apresentacao (reabrir depois que os dados mudaram pergunta de novo). Nunca propaga erro.
Public Sub PptTqq_OnClose(pres As Object)
    On Error Resume Next
    If mSeen Is Nothing Then Exit Sub
    mSeen.Remove LCase$(CStr(pres.FullName))
    Err.Clear
End Sub

' Pura: so trata a apresentacao com janela num PowerPoint visivel (abertura por automacao, sem janela, nunca pergunta nem atualiza).
Public Function PptTqq_ShouldHandle(nWindows As Long, appVisible As Boolean) As Boolean
    PptTqq_ShouldHandle = (nWindows > 0 And appVisible)
End Function

' Eventos (AfterPresentationOpen, WindowActivate) e Auto_Open. Sem janela ou PowerPoint invisivel: nada (WindowActivate cobre a janela
' que aparece depois). Erro: so log (nada de MsgBox no caminho dos eventos).
Public Sub PptTqq_OnOpen(pres As Object)
    Dim nw As Long, vis As Boolean
    On Error GoTo Fail
    If BUILD_CONFIG = "Test" Then mTestOnOpenCalls = mTestOnOpenCalls + 1
    nw = pres.Windows.Count
    vis = (CLng(Application.Visible) <> 0)
    If Not PptTqq_ShouldHandle(nw, vis) Then Exit Sub
    PptTqq_HandleOpen pres
    Exit Sub
Fail:
    PptTqq_Log "PptTqq_OnOpen: " & Err.Number & " " & Err.Description
End Sub

' Decisao ao abrir: a apresentacao e marcada como vista na primeira vez (com ou sem graficos marcados: colar depois nao pergunta) e
' esquecida ao fechar. Pergunta/atualiza conforme PptTqq_OnOpenAction. Reentrada ignorada. Nunca propaga erro.
Public Sub PptTqq_HandleOpen(pres As Object)
    Dim n As Long, act As String, key As String, seen As Boolean, md As String, srcs As Collection, remote As Boolean
    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Fail
    key = LCase$(CStr(pres.FullName))
    seen = PptTqq_Seen(key)
    PptTqq_MarkSeen key
    If Not seen Then
        PptTqq_RelinkData pres, False             ' apresentacao movida: graficos criados pelo PowerPoint voltam a achar os dados (sem tocar a rede)
        Set srcs = New Collection
        n = PptTqq_CountTagged(pres, srcs)
        md = PptTqq_OpenMode()
        If md = "always" And n > 0 Then remote = PptTqq_HasRemoteSource(srcs, PptTqq_ExcelOpenNames())
        act = PptTqq_OnOpenAction(md, (CLng(pres.ReadOnly) <> 0), n, False, remote)
        If act = "sync" Then
            PptTqq_DoSync pres, False
        ElseIf act = "ask" Then
            If PptTqq_Ask(PptTqq_AskText(n, srcs)) = vbYes Then PptTqq_DoSync pres, False
        End If
    End If
    mBusy = False
    Exit Sub
Fail:
    mBusy = False
    PptTqq_Log "Não foi possível verificar os gráficos do ThinQQ desta apresentação: " & Err.Number & " " & Err.Description
End Sub

' ---- 1.6.0: atualizacao automatica por elemento. Indice: COPIA EXATA das PptIdx_* do core (mesmos testes de cadeia). Puras e totais: a
' etiqueta vem da apresentacao (de quem a fez), entao lixo nunca levanta erro, o SlideID so vale com 1 a 9 algarismos, entradas alem de
' PPT_IDX_MAX_ENTRIES sao ignoradas e uma etiqueta maior que PPT_IDX_MAX_LEN e recusada inteira. Entradas por vbLf, campos por "|" (o .pptx
' salvo perde Chr(30)/Chr(31) e troca Tab por espaco: medido no core). Os separadores nunca entram num campo. ----

Public Function PptTqq_IdxSafeField(s As String) As String
    Dim t As String
    t = Replace(Replace(Replace(s, PPT_IDX_FS, "_"), vbLf, "_"), vbCr, "_")
    PptTqq_IdxSafeField = Replace(Replace(Replace(t, vbTab, "_"), Chr$(30), "_"), Chr$(31), "_")
End Function

Public Function PptTqq_IdxEntry(slideId As Long, shapeName As String, book As String) As String
    PptTqq_IdxEntry = CStr(slideId) & PPT_IDX_FS & PptTqq_IdxSafeField(shapeName) & PPT_IDX_FS & PptTqq_IdxSafeField(book)
End Function

Public Function PptTqq_IdxParse(idx As String) As Collection
    Dim out As New Collection, parts() As String, f() As String, i As Long
    Set PptTqq_IdxParse = out
    On Error GoTo Done
    If Len(idx) = 0 Or Len(idx) > PPT_IDX_MAX_LEN Then Exit Function
    parts = Split(idx, PPT_IDX_ES)
    For i = LBound(parts) To UBound(parts)
        If out.Count >= PPT_IDX_MAX_ENTRIES Then Exit For
        If Len(parts(i)) > 0 Then
            f = Split(parts(i), PPT_IDX_FS)
            If UBound(f) = 2 Then
                If Len(f(0)) > 0 And Len(f(0)) < 10 And Len(f(1)) > 0 Then
                    If Not f(0) Like "*[!0-9]*" Then out.Add Array(CLng(f(0)), f(1), f(2))
                End If
            End If
        End If
    Next
Done:
End Function

Public Function PptTqq_IdxJoin(entries As Collection) As String
    Dim e As Variant, a() As String, n As Long
    If entries.Count = 0 Then Exit Function
    ReDim a(0 To entries.Count - 1)
    For Each e In entries
        a(n) = PptTqq_IdxEntry(CLng(e(0)), CStr(e(1)), CStr(e(2)))
        n = n + 1
    Next
    PptTqq_IdxJoin = Join(a, PPT_IDX_ES)
End Function

Public Function PptTqq_IdxSet(idx As String, slideId As Long, shapeName As String, book As String, turnOn As Boolean) As String
    Dim out As New Collection, e As Variant, nm As String
    nm = PptTqq_IdxSafeField(shapeName)
    For Each e In PptTqq_IdxParse(idx)
        If Not (CLng(e(0)) = slideId And StrComp(CStr(e(1)), nm, vbBinaryCompare) = 0) Then out.Add e
    Next
    If turnOn Then out.Add Array(slideId, nm, book)
    PptTqq_IdxSet = PptTqq_IdxJoin(out)
End Function

Private Sub PptTqq_TagDelete(o As Object, key As String)
    On Error Resume Next
    o.Tags.Delete key
    Err.Clear
End Sub

' Medido na C3: Shapes(nome) do PowerPoint acha a forma IGNORANDO a caixa ("tqq_A" e "tqq_a" no slide: os dois nomes acham a primeira), entao
' a comparacao de nomes aqui e sem diferenciar maiusculas (vbTextCompare), como no core.
Private Function PptTqq_SlideHasName(sl As Object, nm As String) As Boolean
    Dim s As Object
    For Each s In sl.Shapes
        If StrComp(CStr(s.Name), nm, vbTextCompare) = 0 Then PptTqq_SlideHasName = True: Exit Function
    Next
End Function

' Copia de PptSync_EnsureUniqueName do core: separadores do indice e caracteres que o .pptx salvo muda viram "_" (PptTqq_IdxSafeField; medido
' na revisao da C1: o PowerPoint ACEITA Chr(30) no nome de uma forma) e nome repetido no slide (sem diferenciar maiusculas; outra forma =
' Shape.Id diferente) vira "_2", "_3"...
Private Sub PptTqq_EnsureUniqueName(shp As Object)
    Dim sl As Object, s As Object, base As String, n As Long, clash As Boolean, myId As Long
    Set sl = shp.Parent
    base = CStr(shp.Name): myId = CLng(shp.Id)
    If PptTqq_IdxSafeField(base) <> base Then
        base = PptTqq_IdxSafeField(base)
        If Not PptTqq_SlideHasName(sl, base) Then shp.Name = base: Exit Sub
        clash = True
    End If
    If Not clash Then
        For Each s In sl.Shapes
            If CLng(s.Id) <> myId And StrComp(CStr(s.Name), base, vbTextCompare) = 0 Then clash = True: Exit For
        Next
    End If
    If Not clash Then Exit Sub
    For n = 2 To 999
        If Not PptTqq_SlideHasName(sl, base & "_" & n) Then shp.Name = base & "_" & n: Exit Sub
    Next
End Sub

' Formas do ThinQQ (TQQ_CHART) da selecao da janela ativa (formas ou texto dentro de uma forma), so as de um slide (mestre/layout nao tem
' SlideID e o indice nao as acha). Vazia sem janela/selecao. Nunca propaga erro (chamada pelos callbacks da faixa a cada mudanca de selecao).
Public Function PptTqq_SelectedTagged() As Collection
    Dim out As New Collection, sel As Object, sr As Object, shp As Object, i As Long, sid As Long
    Set PptTqq_SelectedTagged = out
    On Error Resume Next
    If Application.Windows.Count = 0 Then Exit Function
    Set sel = Application.ActiveWindow.Selection
    If sel Is Nothing Then Exit Function
    If sel.Type <> 2 And sel.Type <> 3 Then Exit Function          ' ppSelectionShapes, ppSelectionText
    Set sr = sel.ShapeRange
    If sr Is Nothing Then Exit Function
    For i = 1 To sr.Count
        Set shp = Nothing
        Set shp = sr(i)
        If Not shp Is Nothing Then
            If Len(PptTqq_Tag(shp, PPT_TAG_CHART)) > 0 Then
                Err.Clear
                sid = CLng(shp.Parent.SlideID)
                If Err.Number = 0 Then out.Add shp
            End If
        End If
        Err.Clear
    Next
    Err.Clear
End Function

' Todas as formas dadas sao automaticas (e ha pelo menos uma).
Public Function PptTqq_ShapesAuto(shapes As Collection) As Boolean
    Dim s As Object
    If shapes.Count = 0 Then Exit Function
    For Each s In shapes
        If PptTqq_Tag(s, PPT_TAG_AUTO) <> "1" Then Exit Function
    Next
    PptTqq_ShapesAuto = True
End Function

' Liga/desliga "Atualizar automaticamente" (etiqueta + indice da apresentacao), sem o Excel. Mesma ordem do core (PptSync_SetAuto): o nome
' fica definitivo antes do indice e a etiqueta da forma (a verdade) e gravada antes do indice (o atalho, reconstruido pela sincronizacao).
Public Sub PptTqq_SetAutoShapes(shapes As Collection, turnOn As Boolean)
    Dim s As Object, pres As Object, idx As String, book As String
    For Each s In shapes
        Set pres = s.Parent.Parent
        If turnOn Then PptTqq_EnsureUniqueName s
        If turnOn Then s.Tags.Add PPT_TAG_AUTO, "1" Else PptTqq_TagDelete s, PPT_TAG_AUTO
        book = PptTqq_SourcePath(PptTqq_LinkSource(s), PptTqq_Tag(s, PPT_TAG_BOOK))
        idx = PptTqq_IdxSet(PptTqq_Tag(pres, PPT_TAG_INDEX), CLng(s.Parent.SlideID), CStr(s.Name), book, turnOn)
        If Len(idx) = 0 Then PptTqq_TagDelete pres, PPT_TAG_INDEX Else pres.Tags.Add PPT_TAG_INDEX, idx
    Next
End Sub

' Entradas (formato do indice, pasta vazia) das formas dadas: o que Core_PptSyncShapes recebe.
Public Function PptTqq_EntriesOf(shapes As Collection) As String
    Dim c As New Collection, s As Object
    For Each s In shapes
        c.Add Array(CLng(s.Parent.SlideID), CStr(s.Name), "")
    Next
    PptTqq_EntriesOf = PptTqq_IdxJoin(c)
End Function

' O core tem Core_PptSyncShapes (API por recurso: a sincronizacao ao abrir so exige PPT_REQUIRED_API). Pura.
Public Function PptTqq_ShapesApiOk(coreApi As Long) As Boolean
    PptTqq_ShapesApiOk = (coreApi >= PPT_API_SHAPES)
End Function

' Atualiza as entradas dadas. Core com API 2 => Core_PptSyncShapes (o texto do core, inclusive o de PowerPoint ocupado, e exibido como veio).
' "Atualizar uma vez" (fromAuto False): core mais velho => a apresentacao inteira com UM aviso (sem Core_PptApiVersion: so os dados e o aviso
' de core antigo; API 1: Core_PptSyncPresentation e o aviso de "so os selecionados"); sem o core => so os dados; sem Excel => cria um.
' "Atualizar automaticamente" (fromAuto, chamado depois de gravar a marca; design secao 8): so o Excel ja aberto (nunca cria um); sem Excel ou
' sem o core => so a dica de abrir o Excel com o ThinQQ; core mais velho => so o aviso de atualizar o ThinQQ (nada e atualizado); API 2 =>
' atualiza e, se a atualizacao automatica esta desligada no Excel, acrescenta onde liga-la. A API e lida uma vez. Reentrada ignorada;
' mBusy liberado em todo caminho antes do aviso. Nunca propaga.
Public Sub PptTqq_SyncEntries(pres As Object, entries As String, force As Boolean, Optional fromAuto As Boolean = False)
    Dim xl As Object, created As Boolean, core As String, msg As String, v As Variant, en As Long, ed As String, api As Long
    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Fail
    Set xl = PptTqq_GetExcel(created, fromAuto)
    If xl Is Nothing Then
        mBusy = False
        If fromAuto Then PptTqq_Msg PPT_MSG_AUTO_NOCORE, vbInformation Else PptTqq_Msg PPT_MSG_NO_EXCEL, vbExclamation
        Exit Sub
    End If
    core = PptTqq_EnsureCore(xl, created)
    If Len(core) = 0 Then
        If fromAuto Then msg = PPT_MSG_AUTO_NOCORE Else msg = PptTqq_DataOnly(pres, PptTqq_WorkbookNames(xl)) & vbLf & PPT_MSG_DATA_ONLY
    Else
        api = PptTqq_CoreApi(xl, core)
        If PptTqq_ShapesApiOk(api) Then
            v = xl.Run("'" & core & "'!Core_PptSyncShapes", pres, entries, force)
            msg = CStr(v)
            If fromAuto Then
                If Not PptTqq_CoreAutoOn(xl, core) Then msg = msg & vbLf & PPT_MSG_AUTO_OFF
            End If
        ElseIf fromAuto Then
            msg = PPT_MSG_AUTO_OLDCORE
        ElseIf PptTqq_ApiCompatible(api) Then
            msg = PptTqq_RunCoreSync(xl, core, pres, force, api) & vbLf & PPT_MSG_SEL_OLDCORE
        Else
            msg = PptTqq_RunCoreSync(xl, core, pres, force, api)          ' so os dados + PPT_MSG_OLD_CORE (um aviso so)
        End If
    End If
    PptTqq_Release xl, created
    mBusy = False
    PptTqq_Msg msg, vbInformation
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptTqq_Release xl, created
    mBusy = False
    PptTqq_Msg PptTqq_ErrorText(en, ed), vbExclamation
End Sub

' Opcao "autoon" do core (API 2). Erro ou Excel ocupado: True (nenhuma nota; o resumo ja saiu). Nunca propaga.
Private Function PptTqq_CoreAutoOn(xl As Object, core As String) As Boolean
    Dim v As Variant
    PptTqq_CoreAutoOn = True
    On Error Resume Next
    v = xl.Run("'" & core & "'!Core_PptGetFlag", "autoon")
    If Err.Number = 0 Then PptTqq_CoreAutoOn = CBool(v)
    Err.Clear
End Function

' Botao "Atualizar automaticamente": grava a marca (e o indice) e, ao ligar, ja atualiza uma vez so essas formas (como o "Update
' Automatically" do think-cell). Desligar nao chama o Excel. Nunca propaga erro (falha ao gravar a marca: aviso proprio).
Public Sub PptTqq_ToggleAuto(shapes As Collection, turnOn As Boolean)
    Dim pres As Object, en As Long, ed As String
    If shapes.Count = 0 Then PptTqq_Msg PPT_MSG_SELECT: Exit Sub
    On Error GoTo Fail
    Set pres = shapes(1).Parent.Parent
    PptTqq_SetAutoShapes shapes, turnOn
    On Error GoTo 0
    If turnOn Then PptTqq_SyncEntries pres, PptTqq_EntriesOf(shapes), True, True
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    If turnOn Then
        PptTqq_Msg "Não foi possível marcar o gráfico como automático: " & ed & " (" & en & ")", vbExclamation
    Else
        PptTqq_Msg "Não foi possível voltar o gráfico para a atualização manual: " & ed & " (" & en & ")", vbExclamation
    End If
End Sub

' Botao "Atualizar uma vez": deixa o nome de cada forma unico no slide ANTES de montar as entradas (um grafico copiado e colado no mesmo
' slide tem o mesmo nome, e a busca do core por nome ignora a caixa: atualizaria a outra forma) e atualiza so essas formas. Nunca propaga.
Public Sub PptTqq_UpdateSelected(shapes As Collection)
    Dim s As Object, en As Long, ed As String
    If shapes.Count = 0 Then PptTqq_Msg PPT_MSG_SELECT: Exit Sub
    On Error GoTo Fail
    For Each s In shapes
        PptTqq_EnsureUniqueName s
    Next
    On Error GoTo 0
    PptTqq_SyncEntries shapes(1).Parent.Parent, PptTqq_EntriesOf(shapes), True
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    PptTqq_Msg PptTqq_ErrorText(en, ed), vbExclamation
End Sub

' ---- 1.8.0 (Parte G): criar graficos pelo PowerPoint. spec 2026-10-15 secao 14. ----

' Copia exata de PptData_PathFor do core (mesmos testes de cadeia). Pura.
Public Function PptTqq_DataPathFor(presFullName As String) As String
    Dim p As Long, k As Long, base As String, f As String
    f = Trim$(presFullName)
    If LCase$(Left$(f, 7)) = "http://" Or LCase$(Left$(f, 8)) = "https://" Then Exit Function
    p = InStrRev(f, "\")
    If p < 2 Then Exit Function
    base = Mid$(f, p + 1)
    k = InStrRev(base, ".")
    If k > 1 Then base = Left$(base, k - 1)
    PptTqq_DataPathFor = Left$(f, p - 1) & "\" & base & PPT_DATA_SUFFIX
End Function

' Slide ativo da apresentacao (modo Normal; senao o selecionado). Nothing em modos de mestre ou sem janela.
Public Function PptTqq_ActiveSlide(pres As Object) As Object
    Dim s As Object, k As Long
    On Error Resume Next
    Set s = Application.ActiveWindow.View.Slide
    k = 0: k = s.SlideIndex
    If k = 0 Then Set s = Application.ActiveWindow.Selection.SlideRange(1): k = s.SlideIndex
    If k = 0 Then Set s = Nothing
    If Not s Is Nothing Then
        If StrComp(CStr(s.Parent.FullName), CStr(pres.FullName), vbTextCompare) <> 0 Then Set s = Nothing
    End If
    Err.Clear
    Set PptTqq_ActiveSlide = s
End Function

Private Sub PptTqq_OpenLoader(xl As Object)
    Dim p As String
    On Error Resume Next
    p = Environ$("LOCALAPPDATA") & "\ThinQQ\ThinQQ.xlam"
    If Len(Dir$(p)) = 0 Then Exit Sub
    xl.Workbooks.Open p, 0, True                    ' somente leitura, sem atualizar vinculos (como o core em PptTqq_TryOpenCore)
    Err.Clear
End Sub

' Confere o Excel obtido para criar/editar dados: core respondendo e com a API 3. Se nao serve, why recebe o aviso (sem o ThinQQ / ThinQQ
' antigo) e, se a instancia foi criada NESTA chamada, ela recebe Quit (nenhum Excel vazio sobra, nada fica em mEditXl). Ocupado: erro.
Private Function PptTqq_UsableEditExcel(xl As Object, created As Boolean, core As String, ByRef why As String) As Object
    Dim api As Long, en As Long, ed As String
    If Len(core) = 0 Then
        why = PPT_MSG_DATA_NOCORE
    Else
        On Error Resume Next
        api = PptTqq_CoreApi(xl, core)
        en = Err.Number: ed = Err.Description
        Err.Clear
        On Error GoTo 0
        If en <> 0 Then
            If created Then PptTqq_Release xl, True
            Err.Raise en, "PptTqq", ed
        End If
        If api < PPT_API_DATA Then why = PPT_MSG_OLD_CORE_DATA
    End If
    If Len(why) > 0 Then
        If created Then PptTqq_Release xl, True
        Exit Function
    End If
    Set PptTqq_UsableEditExcel = xl
End Function

' Excel para criar/editar dados, que FICA com o usuario (spec 14.3): (1) o criado antes por aqui, se ainda esta VISIVEL e responde com o
' core; (2) o Excel aberto COM o core; (3) senao, se ha um core instalado, um novo (created = True): carregado INVISIVEL (alertas desligados,
' loader instalado somente leitura, que carrega core e faixa, ou o core pelo caminho de sempre). Nunca carrega o ThinQQ num Excel do usuario
' que nao o tem. O criado aqui que nao carregou um core com a API 3 recebe Quit na hora; o que deu certo continua INVISIVEL ate o chamador
' conseguir criar/abrir os dados (PptTqq_HandOver: visivel, UserControl, fica com o usuario) ou desistir (PptTqq_Discard: Quit) — numa recusa
' do core nenhum Excel vazio fica a vista (revisao final M2). Nothing: why = o aviso (vazio = sem Excel). Ocupado: erro de ocupado.
' mEditXl (revisao final M1): o usuario fechar o Excel pelo X so esconde a janela enquanto o PowerPoint segura a referencia; instancia
' guardada INVISIVEL = fechada: a referencia e solta (o processo termina) e nunca e reaproveitada.
Public Function PptTqq_GetEditExcel(ByRef core As String, Optional ByRef why As String, Optional ByRef created As Boolean) As Object
    Dim xl As Object, st As Long
    core = "": why = "": created = False
    If BUILD_CONFIG = "Test" Then
        If mTestXl Is Nothing Then Exit Function
        created = mTestCreated
        core = PptTqq_EnsureCore(mTestXl, mTestCreated)
        Set PptTqq_GetEditExcel = PptTqq_UsableEditExcel(mTestXl, mTestCreated, core, why)
        If PptTqq_GetEditExcel Is Nothing Then created = False
        Exit Function
    End If
    PptTqq_ForgetClosedEditExcel
    If Not mEditXl Is Nothing Then
        st = PptTqq_PingState(mEditXl)
        If st = -1 Then PptTqq_RaiseBusy
        If st = 1 Then
            core = PptTqq_CoreFileName()
            Set PptTqq_GetEditExcel = PptTqq_UsableEditExcel(mEditXl, False, core, why)
            Exit Function
        End If
        Set mEditXl = Nothing                                       ' o usuario fechou (ping 0)
    End If
    On Error Resume Next
    Set xl = GetObject(, "Excel.Application")
    Err.Clear
    On Error GoTo 0
    If Not xl Is Nothing Then
        st = PptTqq_PingState(xl)
        If st = -1 Then PptTqq_RaiseBusy
        If st = 1 Then
            core = PptTqq_CoreFileName()
            Set PptTqq_GetEditExcel = PptTqq_UsableEditExcel(xl, False, core, why)
            Exit Function
        End If
    End If
    Set xl = Nothing
    If Len(PptTqq_ListCoreCandidates(PptTqq_CoreDir())) = 0 Then why = PPT_MSG_DATA_NOCORE: Exit Function   ' sem core instalado: nem cria
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    Err.Clear
    On Error GoTo 0
    If xl Is Nothing Then Exit Function
    On Error Resume Next
    xl.DisplayAlerts = False
    Err.Clear
    On Error GoTo 0
    PptTqq_OpenLoader xl
    On Error Resume Next
    If PptTqq_PingState(xl) = 1 Then core = PptTqq_CoreFileName() Else core = PptTqq_EnsureCore(xl, True)
    If Err.Number <> 0 Then core = ""                               ' ocupado logo apos criar: tratado como sem core (Quit abaixo)
    Err.Clear
    On Error GoTo 0
    Set xl = PptTqq_UsableEditExcel(xl, True, core, why)
    If xl Is Nothing Then core = "": Exit Function
    created = True
    Set PptTqq_GetEditExcel = xl
End Function

' mEditXl invisivel (o usuario fechou pelo X) ou morto: solta a referencia. Ocupado: erro de ocupado. Fora do build Test.
Private Sub PptTqq_ForgetClosedEditExcel()
    Dim v As Boolean, en As Long
    If mEditXl Is Nothing Then Exit Sub
    On Error Resume Next
    v = CBool(mEditXl.Visible)
    en = Err.Number
    Err.Clear
    On Error GoTo 0
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Then PptTqq_RaiseBusy
    If en <> 0 Or Not v Then Set mEditXl = Nothing
End Sub

' Excel criado nesta chamada e que deu certo: alertas de volta, visivel, UserControl e guardado (fica com o usuario; sonda PG1). Nunca propaga.
Private Sub PptTqq_HandOver(xl As Object)
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then Exit Sub
    xl.DisplayAlerts = True                                         ' do usuario: "Salvar?" ao fechar volta a aparecer (spec 14.2)
    xl.Visible = True
    xl.UserControl = True
    Set mEditXl = xl
    Err.Clear
End Sub

' Excel criado nesta chamada e que nao serviu (recusa ou falha do core): Quit, nada guardado. Nunca propaga.
Private Sub PptTqq_Discard(xl As Object)
    On Error Resume Next
    If Not mEditXl Is Nothing Then
        If mEditXl Is xl Then Set mEditXl = Nothing
    End If
    PptTqq_Release xl, True
    Err.Clear
End Sub

' Texto de falha de criar/editar dados: ocupado e Excel ausente com os textos de sempre; o resto com o prefixo da acao. Pura.
Private Function PptTqq_DataErrorText(prefix As String, en As Long, ed As String) As String
    If en = RPC_E_CALL_REJECTED Or en = RPC_E_RETRYLATER Then
        PptTqq_DataErrorText = PptTqq_ErrorText(en, ed)
    ElseIf en = 429 Then
        PptTqq_DataErrorText = PPT_MSG_NO_EXCEL_DATA
    Else
        PptTqq_DataErrorText = prefix & ed & " (" & en & ")"
    End If
End Function

' Abre a planilha dos dados da forma no Excel dado e traz a janela a frente (AppActivate com a legenda devolvida pelo core; sonda PG2,
' verificacao manual). "!" do core = a mensagem. Falha: aviso proprio (o grafico ja inserido continua). Devolve True se os dados abriram.
' Excel criado nesta chamada (created): entregue ao usuario (PptTqq_HandOver) antes do AppActivate, ou Quit se os dados nao abriram (e o
' grafico nao foi inserido agora: keepXl False). Nunca propaga.
Private Function PptTqq_EditShape(shp As Object, xl As Object, core As String, Optional created As Boolean, Optional keepXl As Boolean) As Boolean
    Dim cap As String, c As New Collection, en As Long, ed As String
    On Error GoTo Fail
    c.Add shp
    cap = CStr(xl.Run("'" & core & "'!Core_PptEditData", shp.Parent.Parent, PptTqq_EntriesOf(c)))
    If Left$(cap, 1) = "!" Then
        If created Then
            If keepXl Then PptTqq_HandOver xl Else PptTqq_Discard xl
        End If
        PptTqq_Msg Mid$(cap, 2), vbExclamation
        Exit Function
    End If
    PptTqq_EditShape = True
    If created Then PptTqq_HandOver xl
    If BUILD_CONFIG = "Test" Then mLastMsg = "edit:" & cap: mTestTrail = mTestTrail & "edit;": Exit Function
    On Error Resume Next
    AppActivate cap
    If Err.Number <> 0 Then Err.Clear: PptTqq_Msg "O Excel abriu a planilha de dados do gráfico (veja a barra de tarefas)."
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    If created Then
        If keepXl Then PptTqq_HandOver xl Else PptTqq_Discard xl
    End If
    PptTqq_Msg PptTqq_DataErrorText("Não foi possível abrir os dados do gráfico: ", en, ed), vbExclamation
End Function

' Aviso unico por sessao no primeiro grafico criado, ANTES de abrir os dados: com o PowerPoint ainda a frente (depois do AppActivate a caixa
' ficaria atras do Excel e, aberta, o PowerPoint recusaria as chamadas do ciclo automatico: o slide nao acompanharia).
Private Sub PptTqq_FirstHint(pres As Object)
    Dim d As String
    If mHintShown Then Exit Sub
    mHintShown = True
    d = PptTqq_DataPathFor(CStr(pres.FullName))
    If BUILD_CONFIG = "Test" Then mTestTrail = mTestTrail & "hint;"
    PptTqq_Msg "Gráfico criado. Clique em OK e digite os dados na planilha do Excel que abre em seguida: o slide acompanha em cerca de 2 segundos." & vbLf & _
               "Os dados ficam em '" & Mid$(d, InStrRev(d, "\") + 1) & "', ao lado da apresentação: envie os dois arquivos juntos (ou use ThinQQ > Empacotar para enviar)."
    If BUILD_CONFIG = "Test" Then mLastHint = mLastMsg
End Sub

' Ganchos de teste: ordem dos passos de inserir ("hint;edit;") e o texto do aviso de primeira vez.
Public Function PptTqq_TrailForTest() As String
    PptTqq_TrailForTest = mTestTrail
End Function

Public Function PptTqq_LastHintForTest() As String
    PptTqq_LastHintForTest = mLastHint
End Function

' "Inserir grafico" (1.8.0): apresentacao salva (senao a mensagem), slide de destino (o dado; Nothing = um slide novo no fim), Excel para editar
' (fica com o usuario), core com a API 3; o core cria o grafico na pasta de dados e cola no slide; depois abre a planilha dos dados. Nunca propaga.
Public Sub PptTqq_InsertChartIn(pres As Object, slide As Object, kind As String)
    Dim xl As Object, core As String, msg As String, sl As Object, n0 As Long, en As Long, ed As String, why As String, newSl As Boolean
    Dim created As Boolean
    If mBusy Then Exit Sub
    If pres Is Nothing Then PptTqq_Msg PPT_MSG_NO_PRES: Exit Sub      ' Modo de Exibicao Protegido: sem ActivePresentation (revisao final M8)
    On Error GoTo Fail
    If Len(CStr(pres.Path)) = 0 Then PptTqq_Msg PPT_MSG_SAVE_FIRST: Exit Sub
    If Len(PptTqq_DataPathFor(CStr(pres.FullName))) = 0 Then PptTqq_Msg PPT_MSG_WEB_PRES: Exit Sub
    mBusy = True
    Set xl = PptTqq_GetEditExcel(core, why, created)
    If xl Is Nothing Then
        mBusy = False
        If Len(why) = 0 Then why = PPT_MSG_NO_EXCEL_DATA
        PptTqq_Msg why, vbExclamation
        Exit Sub
    End If
    Set sl = slide
    If sl Is Nothing Then Set sl = pres.Slides.Add(pres.Slides.Count + 1, 12): newSl = True
    n0 = sl.Shapes.Count
    msg = CStr(xl.Run("'" & core & "'!Core_PptInsertChart", sl, kind))
    mBusy = False
    If Len(msg) > 0 Then
        If newSl Then PptTqq_DropNewSlide sl, n0                     ' recusa do core: o slide em branco criado aqui sai
        If created Then PptTqq_Discard xl                            ' e o Excel criado agora nao fica vazio a vista
        PptTqq_Msg msg, vbExclamation
        Exit Sub
    End If
    PptRx_InvalidateElem
    If sl.Shapes.Count > n0 Then
        PptTqq_FirstHint pres                                        ' antes de trazer o Excel a frente (revisao G4, I2)
        PptTqq_EditShape sl.Shapes(sl.Shapes.Count), xl, core, created, True
    ElseIf created Then
        PptTqq_HandOver xl
    End If
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    mBusy = False
    If newSl Then PptTqq_DropNewSlide sl, n0
    If created Then PptTqq_AfterFail xl, sl, n0
    PptTqq_Msg PptTqq_DataErrorText("Não foi possível criar o gráfico: ", en, ed), vbExclamation
End Sub

' Falha de "Inserir grafico" num Excel criado nesta chamada: com o grafico ja no slide, o Excel fica com o usuario; sem, Quit. Nunca propaga.
Private Sub PptTqq_AfterFail(xl As Object, sl As Object, n0 As Long)
    Dim n As Long
    On Error Resume Next
    n = -1
    If Not sl Is Nothing Then n = CLng(sl.Shapes.Count)
    If n > n0 Then PptTqq_HandOver xl Else PptTqq_Discard xl
    Err.Clear
End Sub

' Apaga o slide criado por "Inserir grafico" se ele continua so com as n0 formas do layout (nada colado). Nunca propaga.
Private Sub PptTqq_DropNewSlide(sl As Object, n0 As Long)
    On Error Resume Next
    If sl Is Nothing Then Exit Sub
    If CLng(sl.Shapes.Count) <= n0 Then sl.Delete
    Err.Clear
End Sub

' "Editar dados" das formas dadas (a 1a): reaponta antes (apresentacao movida), Excel para editar, Core_PptEditData. Nunca propaga.
Public Sub PptTqq_EditShapes(pres As Object, shapes As Collection)
    Dim xl As Object, core As String, en As Long, ed As String, why As String, created As Boolean
    If mBusy Then Exit Sub
    If shapes.Count = 0 Then PptTqq_Msg PPT_MSG_SELECT: Exit Sub
    On Error GoTo Fail
    PptTqq_RelinkData pres, True                                     ' acao do usuario: pode conferir o caminho antigo remoto
    mBusy = True
    Set xl = PptTqq_GetEditExcel(core, why, created)
    If xl Is Nothing Then
        mBusy = False
        If Len(why) = 0 Then why = PPT_MSG_NO_EXCEL_DATA
        PptTqq_Msg why, vbExclamation
        Exit Sub
    End If
    PptTqq_EditShape shapes(1), xl, core, created, False             ' recusa num Excel criado agora: Quit (revisao final M2)
    mBusy = False
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    mBusy = False
    If created Then PptTqq_Discard xl
    PptTqq_Msg PptTqq_DataErrorText("Não foi possível abrir os dados do gráfico: ", en, ed), vbExclamation
End Sub

' Apresentacao movida/renomeada (spec 14.2): formas TQQ_DATA cuja pasta de dados sumiu sao reapontadas para uma pasta de dados NA PASTA DA
' APRESENTACAO; o indice das automaticas acompanha. Sem Excel, sem perguntar. Devolve quantas formas. Nunca propaga. Regras (revisao G4):
' - candidatos, nesta ordem: (1) o arquivo com o MESMO nome do vinculo atual na pasta da apresentacao (pasta movida); (2) a pasta de dados
'   do nome ATUAL da apresentacao, so se a forma era da pasta marcada na apresentacao (mesmo nome de arquivo que TQQ_DATABOOK: apresentacao e
'   pasta de dados renomeadas juntas). Assim uma forma nunca e reapontada para a pasta de dados de OUTRO deck (Salvar como + inserir cria uma
'   segunda pasta de dados ao lado da primeira).
' - so caminhos LOCAIS antes: nenhum candidato existe (ou o candidato e o proprio vinculo) => nada, sem tocar o caminho antigo (que pode ser
'   \\servidor de outra maquina: o FileExists esperaria o tempo-limite da rede na abertura). O caminho antigo so e conferido quando ha um
'   candidato local diferente (vinculo antigo ainda valido = nao mexe).
' - so pastas criadas pelo ThinQQ: o nome do arquivo do vinculo termina em PPT_DATA_SUFFIX (revisao final M3: uma forma TQQ_DATA de terceiros
'   apontando para "Orcamento.xlsx" nunca e reapontada para o Orcamento.xlsx do usuario na mesma pasta).
' - probeRemote False (abertura, revisao final I1): um caminho antigo REMOTO (\\servidor, http) nunca e testado; com o candidato local
'   presente, reaponta sem tocar a rede (nada de SMB/NTLM para um host escolhido por quem fez o deck, nem tempo-limite da rede ao abrir).
'   True so em "Editar dados" (acao do usuario).
' Sonda PG4: SourceFullName e gravavel (persiste, mantem forma e animacao) e o PowerPoint NAO resolve caminho relativo. LinkFormat.Update com o
' arquivo ausente abre uma caixa modal que trava a chamada: aqui o candidato e conferido (FileExists) ANTES de gravar e nunca ha Update.
Public Function PptTqq_RelinkData(pres As Object, Optional probeRemote As Boolean = False) As Long
    Dim fso As Object, mark As String, markName As String, sl As Object, shp As Object, oldP As String, oldName As String
    Dim cand As String, c As String, dirP As String, idx As String, curData As String, sfx As String
    On Error Resume Next
    sfx = LCase$(PPT_DATA_SUFFIX)
    mark = PptTqq_Tag(pres, PPT_TAG_DATABOOK)
    If Len(mark) = 0 Then Exit Function
    dirP = CStr(pres.Path)
    If Len(dirP) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function
    markName = PptTqq_FileNameOf(mark)
    curData = PptTqq_DataPathFor(CStr(pres.FullName))
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If PptTqq_Tag(shp, PPT_TAG_DATA) = "1" Then
                oldP = PptTqq_SourcePath(PptTqq_LinkSource(shp), PptTqq_Tag(shp, PPT_TAG_BOOK))
                oldName = PptTqq_FileNameOf(oldP)
                cand = ""
                If Len(oldName) > Len(sfx) And LCase$(Right$(oldName, Len(sfx))) = sfx Then
                    c = dirP & "\" & oldName
                    If StrComp(c, oldP, vbTextCompare) <> 0 Then
                        If CBool(fso.FileExists(c)) Then cand = c
                    End If
                    If Len(cand) = 0 And Len(curData) > 0 And StrComp(oldName, markName, vbTextCompare) = 0 Then
                        If StrComp(curData, oldP, vbTextCompare) <> 0 Then
                            If CBool(fso.FileExists(curData)) Then cand = curData
                        End If
                    End If
                End If
                If Len(cand) > 0 Then
                    Err.Clear
                    If probeRemote Or Not PptTqq_IsRemotePath(oldP) Then
                        If BUILD_CONFIG = "Test" Then mTestProbes = mTestProbes + 1
                        If CBool(fso.FileExists(oldP)) Then cand = ""              ' o vinculo antigo ainda vale: nao mexe
                        If Err.Number <> 0 Then cand = ""
                    End If
                    Err.Clear
                End If
                If Len(cand) > 0 Then
                    shp.LinkFormat.SourceFullName = cand
                    If Err.Number = 0 Then
                        shp.Tags.Add PPT_TAG_BOOK, cand
                        If StrComp(oldName, markName, vbTextCompare) = 0 Then pres.Tags.Add PPT_TAG_DATABOOK, cand
                        If PptTqq_Tag(shp, PPT_TAG_AUTO) = "1" Then
                            idx = PptTqq_IdxSet(PptTqq_Tag(pres, PPT_TAG_INDEX), CLng(sl.SlideID), CStr(shp.Name), cand, True)
                            pres.Tags.Add PPT_TAG_INDEX, idx
                        End If
                        PptTqq_RelinkData = PptTqq_RelinkData + 1
                    Else
                        PptTqq_Log "PptTqq_RelinkData: " & Err.Number & " " & Err.Description
                    End If
                    Err.Clear
                End If
            End If
        Next
    Next
    Err.Clear
End Function

' Nome do arquivo de um caminho (depois da ultima barra \ ou /); "" sem nome. Pura.
Public Function PptTqq_FileNameOf(p As String) As String
    Dim t As String, k As Long
    t = Trim$(p)
    k = InStrRev(Replace(t, "/", "\"), "\")
    PptTqq_FileNameOf = Mid$(t, k + 1)
End Function

' Evento PresentationSave: salva a pasta de dados aberta e alterada no Excel com o ThinQQ (o de edicao, senao o aberto); nunca abre nem cria
' um Excel; nunca mostra mensagem. Sonda PG5: o evento chega DEPOIS da gravacao do .pptx (arquivo ja novo, Saved = True; o BeforeSave vem
' antes) e, no Salvar como, ja com o nome novo; para a pasta de dados separada a ordem nao importa. Nunca propaga.
Public Sub PptTqq_OnSave(pres As Object)
    Dim xl As Object
    On Error Resume Next
    If Len(PptTqq_Tag(pres, PPT_TAG_DATABOOK)) = 0 Then Exit Sub
    If BUILD_CONFIG = "Test" Then
        Set xl = mTestXl
    Else
        PptTqq_ForgetClosedEditExcel                               ' fechado pelo X (invisivel): solta a referencia
        Err.Clear
        If Not mEditXl Is Nothing Then
            Select Case PptTqq_PingState(mEditXl)
                Case 1: Set xl = mEditXl
                Case -1: Exit Sub                                  ' ocupado: nao salva agora (o ciclo automatico salva depois)
                Case Else: Set mEditXl = Nothing                   ' fechado pelo usuario: o aberto
            End Select
        End If
        If xl Is Nothing Then Set xl = GetObject(, "Excel.Application")
    End If
    Err.Clear
    If xl Is Nothing Then Exit Sub
    If PptTqq_PingState(xl) <> 1 Then Exit Sub
    If PptTqq_CoreApi(xl, PptTqq_CoreFileName()) < PPT_API_DATA Then Exit Sub
    xl.Run "'" & PptTqq_CoreFileName() & "'!Core_PptDataSave", pres
    Err.Clear
End Sub

' ---- 1.9.0 (Parte G-II, G9): "Empacotar para enviar" (spec 14.2 e 14.5). Sem Excel, sem ZIP, sem processo filho: so FileSystemObject. ----

' Pastas de dados da apresentacao (revisao G4, M8: Salvar como + inserir deixa formas em DUAS pastas de dados): a marca TQQ_DATABOOK e os
' caminhos das formas TQQ_DATA (vinculo, senao TQQ_BOOK), distintos sem diferenciar maiusculas, so os de nome "<...> - dados ThinQQ.xlsx"
' (nunca uma pasta do usuario). So le etiquetas e vinculos (nunca o disco); nao desce em grupos (como PptTqq_RelinkData). others: formas do
' ThinQQ sem TQQ_DATA (coladas do Excel: a pasta de origem e do usuario e nao entra no pacote). Nunca propaga.
Public Function PptTqq_DataBooksOf(pres As Object, Optional ByRef others As Long) As Collection
    Dim out As New Collection, sl As Object, shp As Object
    Set PptTqq_DataBooksOf = out
    others = 0
    On Error Resume Next
    PptTqq_AddDataBook out, PptTqq_Tag(pres, PPT_TAG_DATABOOK)
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If PptTqq_Tag(shp, PPT_TAG_DATA) = "1" Then
                PptTqq_AddDataBook out, PptTqq_SourcePath(PptTqq_LinkSource(shp), PptTqq_Tag(shp, PPT_TAG_BOOK))
            ElseIf Len(PptTqq_Tag(shp, PPT_TAG_CHART)) > 0 Then
                others = others + 1
            End If
        Next
    Next
    Err.Clear
End Function

Private Sub PptTqq_AddDataBook(out As Collection, p As String)
    Dim t As String, nm As String, sfx As String
    t = Trim$(p)
    nm = PptTqq_FileNameOf(t)
    sfx = LCase$(PPT_DATA_SUFFIX)
    If Len(nm) <= Len(sfx) Or Len(nm) = Len(t) Then Exit Sub          ' sem nome proprio ou sem pasta (caminho relativo)
    If LCase$(Right$(nm, Len(sfx))) <> sfx Then Exit Sub
    On Error Resume Next
    out.Add t, LCase$(t)                                               ' repetida: chave ja existe, ignorada
    Err.Clear
End Sub

Private Function PptTqq_InKeys(c As Collection, key As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = c(key)
    PptTqq_InKeys = (Err.Number = 0)
    Err.Clear
End Function

' Pastas abertas COM alteracoes nao salvas no Excel ja aberto (o de edicao guardado, senao o do GetObject; nunca cria um Excel; build Test: o
' injetado), chave = FullName em minusculas. So le FullName/Saved (nunca abre nem salva nada). Ocupado ou ausente: vazia. Nunca propaga.
Private Function PptTqq_DirtyOpenBooks() As Collection
    Dim out As New Collection, xl As Object, wb As Object
    Set PptTqq_DirtyOpenBooks = out
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then
        Set xl = mTestXl
    Else
        If Not mEditXl Is Nothing Then
            If CBool(mEditXl.Visible) Then Set xl = mEditXl
        End If
        Err.Clear
        If xl Is Nothing Then Set xl = GetObject(, "Excel.Application")
    End If
    Err.Clear
    If xl Is Nothing Then Exit Function
    For Each wb In xl.Workbooks
        If Not CBool(wb.Saved) Then out.Add 1, LCase$(CStr(wb.FullName))
        Err.Clear
    Next
    Err.Clear
End Function

Public Function PptTqq_PackProbesForTest() As Long
    PptTqq_PackProbesForTest = mTestPackProbes
End Function

' Acrescenta um nome (limpo, PptTqq_CleanPath) a uma lista "a, b". Pura.
Private Function PptTqq_ListAdd(lst As String, nm As String) As String
    If Len(lst) > 0 Then PptTqq_ListAdd = lst & ", " & PptTqq_CleanPath(nm) Else PptTqq_ListAdd = PptTqq_CleanPath(nm)
End Function

' "Empacotar para enviar": salva a apresentacao (se nao for somente leitura) e a pasta de dados marcada aberta e alterada (PptTqq_OnSave: so
' com o Excel do ThinQQ ja aberto; nunca abre nem salva uma pasta do usuario) e COPIA a apresentacao e as pastas de dados dela
' (PptTqq_DataBooksOf) para "<pasta da apresentacao>\<nome> (ThinQQ)\" (lugar fixo; criada se nao existe; arquivos de mesmo nome
' substituidos). So entram as pastas de dados NA PASTA da apresentacao (revisao final G-II, I1); as de fora (local ou rede: o caminho vem da
' apresentacao, que pode ser de outra pessoa) nunca sao tocadas e o aviso as cita em linha propria, como as ausentes; o "Pronto" lista cada
' arquivo copiado; uma pasta de dados copiada que esta aberta no Excel com alteracoes nao salvas e citada (foi a versao do disco); uma que nao
' copia (bloqueada pelo Excel: PG10) pede para fecha-la no Excel. Devolve a pasta do pacote ("" se a apresentacao nao foi copiada; o aviso
' diz por que). Reentrada ignorada. Nunca propaga.
Public Function PptTqq_Package(pres As Object) As String
    Dim fso As Object, dest As String, b As Variant, nm As String, done As New Collection, n As Long, others As Long, en As Long, ed As String
    Dim missing As String, outside As String, failed As String, msg As String, roUnsaved As Boolean, copied As String, dirty As String
    Dim openDirty As Collection
    If mBusy Then Exit Function
    If pres Is Nothing Then PptTqq_Msg PPT_MSG_NO_PRES: Exit Function
    On Error GoTo Fail
    If Len(CStr(pres.Path)) = 0 Or Len(PptTqq_DataPathFor(CStr(pres.FullName))) = 0 Then PptTqq_Msg PPT_MSG_PACK_SAVE_FIRST: Exit Function
    mBusy = True
    If CLng(pres.ReadOnly) <> 0 Then
        roUnsaved = (CLng(pres.Saved) = 0)
    ElseIf CLng(pres.Saved) = 0 Then
        pres.Save
    End If
    PptTqq_OnSave pres
    Set fso = CreateObject("Scripting.FileSystemObject")
    dest = CStr(pres.Path) & "\" & CStr(fso.GetBaseName(CStr(pres.FullName))) & PPT_PACK_SUFFIX
    If Not CBool(fso.FolderExists(dest)) Then fso.CreateFolder dest
    fso.CopyFile CStr(pres.FullName), dest & "\", True
    copied = PptTqq_ListAdd("", CStr(fso.GetFileName(CStr(pres.FullName))))
    PptTqq_Package = dest
    Set openDirty = PptTqq_DirtyOpenBooks()
    For Each b In PptTqq_DataBooksOf(pres, others)
        nm = PptTqq_FileNameOf(CStr(b))
        ' revisao final G-II (I1): so as pastas de dados NA PASTA da apresentacao entram no pacote. O caminho vem do deck (pode ser de outra
        ' pessoa e apontar para a pasta de dados confidencial de outra apresentacao): fora dela (local ou rede) nada e copiado nem sondado
        ' no disco, e o aviso a cita em linha propria
        If StrComp(Left$(CStr(b), Len(CStr(b)) - Len(nm) - 1), CStr(pres.Path), vbTextCompare) <> 0 Then
            outside = PptTqq_ListAdd(outside, nm)
        ElseIf PptTqq_InKeys(done, LCase$(nm)) Then
            ' a mesma pasta escrita de outro jeito (nao acontece com caminhos normalizados): ja copiada
        Else
            If BUILD_CONFIG = "Test" Then mTestPackProbes = mTestPackProbes + 1
            If Not PptTqq_FileExists(CStr(b)) Then
                missing = PptTqq_ListAdd(missing, nm)
            Else
                Err.Clear
                On Error Resume Next
                fso.CopyFile CStr(b), dest & "\" & nm, True
                en = Err.Number
                Err.Clear
                On Error GoTo Fail
                If en = 0 Then
                    n = n + 1
                    done.Add 1, LCase$(nm)
                    copied = PptTqq_ListAdd(copied, nm)
                    If PptTqq_InKeys(openDirty, LCase$(CStr(b))) Then dirty = PptTqq_ListAdd(dirty, nm)
                Else
                    failed = PptTqq_ListAdd(failed, nm)
                End If
            End If
        End If
    Next
    If Len(copied) > 300 Then copied = Left$(copied, 297) & "..."
    msg = "Pronto: a apresentação e " & n & " pasta(s) de dados do ThinQQ foram copiadas para '" & PptTqq_CleanPath(dest) & "': " & copied & ". Envie a pasta " & _
          "inteira (por exemplo, compactada): quem abrir a apresentação dessa pasta com o ThinQQ terá os gráficos criados pelo PowerPoint ligados aos dados ao lado."
    ' as linhas de excecao mais importantes primeiro (o aviso e cortado em PPT_MAX_ASK_LEN)
    If Len(outside) > 0 Then msg = msg & vbLf & "Não copiada(s): fora da pasta da apresentação: " & outside
    If Len(dirty) > 0 Then msg = msg & vbLf & "Copiada(s) a versão salva no disco (aberta(s) no Excel com alterações não salvas; salve-a(s) no Excel e empacote de novo para levá-las): " & dirty
    If Len(failed) > 0 Then msg = msg & vbLf & "Não foi possível copiar (feche-a no Excel e empacote de novo): " & failed
    If Len(missing) > 0 Then msg = msg & vbLf & "Não encontrada(s) (os gráficos ficam com os últimos dados): " & missing
    If roUnsaved Then msg = msg & vbLf & "A apresentação está somente leitura: foi copiada a versão salva no disco (sem as alterações abertas)."
    If others > 0 Then msg = msg & vbLf & "Gráficos colados do Excel continuam ligados às pastas de origem, que não entram no pacote."
    If Len(msg) > PPT_MAX_ASK_LEN Then msg = Left$(msg, PPT_MAX_ASK_LEN - 3) & "..."
    mBusy = False
    If Len(failed) + Len(missing) + Len(outside) + Len(dirty) > 0 Or roUnsaved Then PptTqq_Msg msg, vbExclamation Else PptTqq_Msg msg, vbInformation
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    mBusy = False
    PptTqq_Package = ""
    PptTqq_Msg "Não foi possível empacotar: " & ed & " (" & en & ")", vbExclamation
End Function

' ---- 1.9.0 (Parte G-II, G8): grupo "Grafico do ThinQQ" (spec 14.5). O estado da faixa vem da etiqueta TQQ_OPTS (sem chamar o Excel). ----

Public Sub PptTqq_SetPickForTest(a As Long, b As Long)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestPickA = a: mTestPickB = b
End Sub

' Valor de uma chave em TQQ_OPTS ("a=1;b=x;caps=..."); "" sem a chave. Pura.
Public Function PptTqq_OptValue(opts As String, key As String) As String
    Dim p As Variant, k As Long
    If Len(key) = 0 Then Exit Function
    For Each p In Split(opts, ";")
        k = InStr(CStr(p), "=")
        If k > 1 Then
            If Left$(CStr(p), k - 1) = key Then PptTqq_OptValue = Mid$(CStr(p), k + 1): Exit Function
        End If
    Next
End Function

' O recurso esta na lista caps de TQQ_OPTS (nome inteiro: "tal" nao casa com "totals"). Pura.
Public Function PptTqq_OptCap(opts As String, feature As String) As Boolean
    If Len(feature) = 0 Then Exit Function
    PptTqq_OptCap = InStr("," & PptTqq_OptValue(opts, "caps"), "," & feature & ",") > 0
End Function

' Item selecionado de uma lista do grupo: tag + valor quando o valor e um dos itens da faixa; senao (sem selecao, legenda "lineend"/"inbar"
' definida no Excel, valor estranho) o 1o item, sempre um id que existe (revisao final G-II, M5). Pura.
Public Function PptTqq_OptItemId(tag As String, v As String) As String
    Dim items As String
    Select Case tag
        Case "numfmt": items = ",int,dec1,dec2,pct0,pct1,mil,mi,"
        Case "legend": items = ",right,none,"
        Case Else: Exit Function
    End Select
    If Len(v) > 0 And InStr(items, "," & v & ",") > 0 Then
        PptTqq_OptItemId = tag & v
    Else
        PptTqq_OptItemId = tag & Split(items, ",")(1)
    End If
End Function

' Gancho de teste: PptRx_InvalidateElem chamado (o projeto de teste nao tem faixa).
Public Sub PptTqq_NoteInvalidate()
    If BUILD_CONFIG = "Test" Then mTestInvalidates = mTestInvalidates + 1
End Sub

Public Function PptTqq_InvalidatesForTest() As Long
    PptTqq_InvalidatesForTest = mTestInvalidates
End Function

' TQQ_OPTS da 1a forma do ThinQQ selecionada ("" sem selecao ou forma colada por um core anterior ao 1.9.0, ate a proxima atualizacao).
Public Function PptTqq_SelectedOpts() As String
    Dim sel As Collection
    On Error Resume Next
    Set sel = PptTqq_SelectedTagged()
    If sel.Count > 0 Then PptTqq_SelectedOpts = PptTqq_Tag(sel(1), PPT_TAG_OPTS)
    Err.Clear
End Function

' Numero de categoria digitado (so algarismos, ate 6); -1 se invalido. Pura.
Private Function PptTqq_PickNumber(v As String) As Long
    Dim t As String
    t = Trim$(v)
    PptTqq_PickNumber = -1
    If Len(t) = 0 Or Len(t) > 6 Then Exit Function
    If t Like "*[!0-9]*" Then Exit Function
    PptTqq_PickNumber = CLng(t)
End Function

' Escolha de/ate pela lista numerada das categorias (cats: nomes separados por Chr(31), valor devolvido pelo Core_PptCategories; so na
' memoria, nunca gravado numa etiqueta). InputBox do PowerPoint. True so com 1 <= de < ate <= n. Cancelado: a = b = 0; resposta invalida:
' a e b diferentes de 0 (o chamador avisa). No build Test responde com o gancho (PptTqq_SetPickForTest).
Public Function PptTqq_PickTwo(cats As String, ByRef a As Long, ByRef b As Long) As Boolean
    Dim c() As String, i As Long, lst As String, n As Long, v As String
    a = 0: b = 0
    c = Split(cats, Chr$(31)): n = UBound(c) + 1
    If n < 2 Then Exit Function
    If BUILD_CONFIG = "Test" Then
        a = mTestPickA: b = mTestPickB
    Else
        For i = 0 To n - 1
            lst = lst & (i + 1) & " - " & Left$(c(i), 60) & vbLf
            If Len(lst) > 600 Then lst = lst & "..." & vbLf: Exit For
        Next
        v = InputBox(lst & vbLf & "Da categoria número:", PPT_APP_NAME, "1")
        If Len(v) = 0 Then Exit Function
        a = PptTqq_PickNumber(v)
        If a < 0 Then b = -1: Exit Function
        v = InputBox(lst & vbLf & "Até a categoria número:", PPT_APP_NAME, CStr(n))
        If Len(v) = 0 Then a = 0: Exit Function
        b = PptTqq_PickNumber(v)
    End If
    PptTqq_PickTwo = (a >= 1 And b <= n And a < b)
End Function

' Excel para um comando: o de edicao guardado (visivel, com o core: a pasta de dados aberta nele e usada como esta) ou o de sempre
' (PptTqq_GetExcel: o aberto com o core, senao uma instancia nova e invisivel, fechada no fim). Ocupado: erro de ocupado.
Private Function PptTqq_CmdExcel(ByRef created As Boolean) As Object
    Dim st As Long
    created = False
    If BUILD_CONFIG <> "Test" Then
        PptTqq_ForgetClosedEditExcel
        If Not mEditXl Is Nothing Then
            st = PptTqq_PingState(mEditXl)
            If st = -1 Then PptTqq_RaiseBusy
            If st = 1 Then Set PptTqq_CmdExcel = mEditXl: Exit Function
        End If
    End If
    Set PptTqq_CmdExcel = PptTqq_GetExcel(created)
End Function

' Comando do grupo "Grafico do ThinQQ" nas formas dadas (spec 14.5): Core_PptCommand com a API 4 (core antigo: PPT_MSG_OLD_CORE_CMD). CAGR e
' diferenca: categorias do 1o grafico (Core_PptCategories), escolha de/ate e "de,ate"; sem categorias (pasta indisponivel) o core e chamado
' com "0,0" so para devolver o motivo (pasta do usuario fechada, pasta de dados ausente ou somente leitura, recusa da matriz). Os nomes das
' formas ficam unicos no slide antes de montar as entradas (como "Atualizar uma vez"). O texto do core (resumo, ocupado, recusa) e exibido
' pelo PowerPoint, depois de soltar o Excel. Reentrada ignorada; mBusy liberado em todo caminho. Nunca propaga.
Public Sub PptTqq_CommandShapes(pres As Object, shapes As Collection, cmd As String, arg As String)
    Dim xl As Object, created As Boolean, core As String, msg As String, cats As String, a As Long, b As Long, a2 As String
    Dim s As Object, entries As String, en As Long, ed As String
    If mBusy Then Exit Sub
    If shapes Is Nothing Then PptTqq_Msg PPT_MSG_SELECT: Exit Sub
    If shapes.Count = 0 Then PptTqq_Msg PPT_MSG_SELECT: Exit Sub
    mBusy = True
    On Error GoTo Fail
    If pres Is Nothing Then Set pres = shapes(1).Parent.Parent
    For Each s In shapes
        PptTqq_EnsureUniqueName s
    Next
    entries = PptTqq_EntriesOf(shapes)
    Set xl = PptTqq_CmdExcel(created)
    If xl Is Nothing Then mBusy = False: PptTqq_Msg PPT_MSG_NO_EXCEL_DATA, vbExclamation: PptRx_InvalidateElem: Exit Sub
    core = PptTqq_EnsureCore(xl, created)
    a2 = arg
    If Len(core) = 0 Then
        msg = PPT_MSG_DATA_NOCORE
    ElseIf PptTqq_CoreApi(xl, core) < PPT_API_CMD Then
        msg = PPT_MSG_OLD_CORE_CMD
    Else
        If cmd = "cagr" Or cmd = "diff" Then
            cats = CStr(xl.Run("'" & core & "'!Core_PptCategories", pres, entries))
            If Len(cats) = 0 Then
                a2 = "0,0"                                          ' so o motivo: o core recusa antes de mudar qualquer coisa
            ElseIf UBound(Split(cats, Chr$(31))) < 1 Then
                msg = PPT_MSG_CATS: cmd = ""
            ElseIf PptTqq_PickTwo(cats, a, b) Then
                a2 = a & "," & b
            Else
                If a <> 0 And b <> 0 Then msg = PPT_MSG_CATS        ' resposta invalida (cancelado: nada)
                cmd = ""
            End If
        End If
        If Len(cmd) > 0 Then
            msg = CStr(xl.Run("'" & core & "'!Core_PptCommand", pres, entries, cmd, a2))
            If a2 = "0,0" And Len(msg) = 0 Then msg = PPT_MSG_CATS
        End If
    End If
    PptTqq_Release xl, created
    mBusy = False
    If Len(msg) > 0 Then PptTqq_Msg msg, vbExclamation
    PptRx_InvalidateElem
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptTqq_Release xl, created
    mBusy = False
    PptTqq_Msg PptTqq_DataErrorText("Não foi possível aplicar: ", en, ed), vbExclamation
    PptRx_InvalidateElem                                ' o botao clicado volta ao estado gravado (revisao final G-II, M3)
End Sub

' ---- Texto vinculado ao Excel (1.9.0, API 5; spec 5.3) ----

' Selecao da janela ativa: texto selecionado ou cursor (ppSelectionText: o trecho) ou UMA forma com texto (o texto inteiro). False se nao
' houver (sem janela, nada selecionado, varias formas, forma sem texto). Selecao dentro de um grupo: False e inGroup = True (o texto vinculado
' so vale para formas soltas: a sincronizacao nao desce em grupos). Nunca propaga erro.
Public Function PptTqq_SelectionTextRange(ByRef shp As Object, ByRef st As Long, ByRef ln As Long, Optional ByRef inGroup As Boolean) As Boolean
    Dim sel As Object, s As Object
    inGroup = False
    On Error GoTo Done
    Set sel = Application.ActiveWindow.Selection
    If sel.Type = 2 Or sel.Type = 3 Then
        If PptTqq_SelInGroup(sel) Then inGroup = True: Exit Function
    End If
    If sel.Type = 3 Then                                         ' ppSelectionText
        Set s = sel.ShapeRange(1)
        st = CLng(sel.TextRange.Start): ln = CLng(sel.TextRange.Length)
    ElseIf sel.Type = 2 Then                                     ' ppSelectionShapes
        If sel.ShapeRange.Count <> 1 Then Exit Function
        Set s = sel.ShapeRange(1)
        If Not CBool(s.HasTextFrame) Then Exit Function
        st = 1: ln = Len(CStr(s.TextFrame.TextRange.Text))
    Else
        Exit Function
    End If
    Set shp = s
    PptTqq_SelectionTextRange = True
Done:
End Function

' Selecao dentro de um grupo (HasChildShapeRange). Leitura que falha = nao. Nunca propaga erro.
Public Function PptTqq_SelInGroup(sel As Object) As Boolean
    On Error Resume Next
    PptTqq_SelInGroup = CBool(sel.HasChildShapeRange)
    Err.Clear
End Function

' Chama Core_PptLinkText no Excel dado, so com a API 5 (core antigo: PPT_MSG_OLD_CORE_TEXT, sem chamar). Devolve o texto do core ("" =
' vinculado). Excel ocupado: o erro de ocupado sobe; outra falha da chamada sobe.
Public Function PptTqq_LinkTextTo(xl As Object, core As String, shp As Object, st As Long, ln As Long) As String
    If PptTqq_CoreApi(xl, core) < PPT_API_TEXT Then PptTqq_LinkTextTo = PPT_MSG_OLD_CORE_TEXT: Exit Function
    PptTqq_LinkTextTo = CStr(xl.Run("'" & core & "'!Core_PptLinkText", shp, st, ln))
End Function

' Aviso final do vincular: "" (vinculado) ou o "vinculado, cortado" do core ganham o lembrete de salvar a pasta (o vincular criou um nome
' oculto nela e a deixou "nao salva"; sem salvar, o vinculo segue pelo endereco e nao acompanha linhas inseridas). Recusas: como vieram. Pura.
Public Function PptTqq_LinkTextMsg(coreMsg As String) As String
    If Len(coreMsg) = 0 Then
        PptTqq_LinkTextMsg = PPT_MSG_TEXT_OK & " " & PPT_MSG_TEXT_SAVE
    ElseIf Left$(coreMsg, 15) = "Texto vinculado" Then              ' "Texto vinculado à célula Pasta!Plan!B2." (e o cortado)
        PptTqq_LinkTextMsg = coreMsg & " " & PPT_MSG_TEXT_SAVE
    Else
        PptTqq_LinkTextMsg = coreMsg
    End If
End Function

' Vincula o trecho [st, st + ln) da forma a celula ATIVA do Excel JA aberto com o ThinQQ (a origem e a celula dele). Nunca cria um Excel nem
' abre o core (sem eles: PPT_MSG_OPEN_EXCEL); o Excel do usuario nunca recebe Quit. O core roda em modo silencioso (sem caixas no Excel) e o
' aviso (sucesso, recusa, core antigo, ocupado) e mostrado pelo PowerPoint depois de a chamada voltar. Reentrada ignorada; mBusy liberado em
' todo caminho. Nunca propaga.
Public Sub PptTqq_LinkText(shp As Object, st As Long, ln As Long)
    Dim xl As Object, created As Boolean, core As String, msg As String, en As Long, ed As String
    If mBusy Then Exit Sub
    If shp Is Nothing Then PptTqq_Msg PPTTEXT_MSG_NOTEXT_PPT: Exit Sub
    mBusy = True
    On Error GoTo Fail
    Set xl = PptTqq_GetExcel(created, True)
    If xl Is Nothing Then
        msg = PPT_MSG_OPEN_EXCEL
    Else
        core = PptTqq_EnsureCore(xl, False)
        If Len(core) = 0 Then
            msg = PPT_MSG_OPEN_EXCEL
        Else
            msg = PptTqq_LinkTextMsg(PptTqq_LinkTextTo(xl, core, shp, st, ln))
        End If
    End If
    PptTqq_Release xl, created
    mBusy = False
    PptTqq_Msg msg
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptTqq_Release xl, created
    mBusy = False
    PptTqq_Msg PptTqq_DataErrorText("Não foi possível vincular o texto: ", en, ed), vbExclamation
End Sub
