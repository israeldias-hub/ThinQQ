Attribute VB_Name = "modPptLinks"
Option Explicit
' Logica do dialogo Vinculos (1.11.0; spec 2026-10-23 secao 8.2). Listar e filtrar so leem etiquetas (em processo, sem Excel); os estados que
' nao estao na etiqueta (pasta nao encontrada, nao encontrado na pasta, sem vinculo, fechada) vem do core (Verificar; abertura; ao abrir o
' dialogo, so do Excel ja aberto) e ficam lembrados na sessao (mStates, por apresentacao). Acoes reaproveitam o .ppam: Atualizar =
' PptTqq_SyncEntriesMsg; Automatico/Manual = PptTqq_SetAutoShapes; Abrir no Excel = PptTqq_EditShapes. Desvincular tira as etiquetas TQQ_*
' (o vinculo nativo do PowerPoint fica: nunca BreakLink, que pode derrubar o PowerPoint; S10). Com a pasta aberta, o vinculo do proprio
' Office pode mover os valores das series sozinho (medido na U1): nenhum texto aqui promete dados congelados.
' O formulario e NAO modal: o evento de selecao (PptLinks_OnSelChange) so o atualiza se ele estiver aberto, nunca o carrega e nunca chama
' o Excel; fechar a apresentacao (PptLinks_OnClose) ou o PowerPoint (PptLinks_CloseAll, no Auto_Close) descarrega o dialogo dela.
Public Const PPTLINKS_MSG_SELECT As String = "Selecione um ou mais elementos na lista."
Public Const PPTLINKS_MSG_EDIT_CHART As String = "Abrir no Excel vale para gráficos. Para mudar um texto ou uma tabela vinculados, altere as células no Excel e clique em Atualizar."
Private Const PPTLINKS_FORMNAME As String = "frmTqqLinks"
Private Const PPTLINKS_TAGS As String = "TQQ_CHART,TQQ_BOOK,TQQ_SHEET,TQQ_VER,TQQ_FP,TQQ_TITLE,TQQ_KIND,TQQ_OPTS,TQQ_DATA,TQQ_TEXT,TQQ_TNAME,TQQ_TREF,TQQ_TPRE,TQQ_TSUF,TQQ_TLAST,TQQ_TABLE,TQQ_TBLAST,TQQ_STATE,TQQ_AUTO"
Private mStates As Object                       ' "apresentacao|slideId|nome em minusculas" -> estado vindo do core nesta sessao
Private mOpen As Boolean                        ' o dialogo foi mostrado e nao fechou (atalho do evento de selecao)

Public Sub PptLinks_ResetTestHooks()
    Set mStates = Nothing
End Sub

Public Function PptLinks_StateLabel(code As String) As String
    Select Case code
        Case "new": PptLinks_StateLabel = "Dados novos"
        Case "ok": PptLinks_StateLabel = "Atualizado"
        Case "missing": PptLinks_StateLabel = "Pasta não encontrada"
        Case "nosource": PptLinks_StateLabel = "Não encontrado na pasta"
        Case "unlinked": PptLinks_StateLabel = "Sem vínculo"
        Case "unknown": PptLinks_StateLabel = "Não verificado (pasta fechada)"
        Case Else: PptLinks_StateLabel = "Não verificado"
    End Select
End Function

Public Function PptLinks_KindLabel(kind As String) As String
    Select Case kind
        Case "text": PptLinks_KindLabel = "Texto"
        Case "table": PptLinks_KindLabel = "Tabela"
        Case Else: PptLinks_KindLabel = "Gráfico"
    End Select
End Function

Public Function PptLinks_ModeLabel(mode As String) As String
    If mode = "auto" Then PptLinks_ModeLabel = "Automático" Else PptLinks_ModeLabel = "Manual"
End Function

Public Function PptLinks_FileName(p As String) As String
    Dim k As Long
    k = InStrRev(p, "\")
    If InStrRev(p, "/") > k Then k = InStrRev(p, "/")
    PptLinks_FileName = Mid$(p, k + 1)
End Function

Public Function PptLinks_OpenModeIndex(mode As String) As Long
    Select Case mode
        Case "ask": PptLinks_OpenModeIndex = 1
        Case "always": PptLinks_OpenModeIndex = 2
        Case Else: PptLinks_OpenModeIndex = 0
    End Select
End Function

Public Function PptLinks_OpenModeOf(index As Long) As String
    Select Case index
        Case 1: PptLinks_OpenModeOf = "ask"
        Case 2: PptLinks_OpenModeOf = "always"
        Case Else: PptLinks_OpenModeOf = "never"
    End Select
End Function

' Chave da apresentacao no cache (FullName em minusculas; "" sem apresentacao). Nunca propaga.
Public Function PptLinks_PresKey(pres As Object) As String
    On Error Resume Next
    If pres Is Nothing Then Exit Function
    PptLinks_PresKey = LCase$(CStr(pres.FullName))
    Err.Clear
End Function

' Chave de uma linha do dialogo (e de uma forma): "SlideID|nome em minusculas" (o PowerPoint acha a forma pelo nome sem diferenciar
' maiusculas). Pura.
Public Function PptLinks_RowKey(r As Variant) As String
    PptLinks_RowKey = CStr(r(1)) & "|" & LCase$(CStr(r(2)))
End Function

Private Function PptLinks_Key(presKey As String, slideId As Variant, nm As Variant) As String
    PptLinks_Key = presKey & "|" & CStr(slideId) & "|" & LCase$(CStr(nm))
End Function

Private Sub PptLinks_EnsureStates()
    If mStates Is Nothing Then Set mStates = CreateObject("Scripting.Dictionary")
End Sub

' Guarda os estados devolvidos pelo core (linhas de Core_PptLinkStates) da apresentacao dada (sem ela: chave sem apresentacao, achada por
' qualquer uma). Devolve quantas linhas validas. "!" (aviso do core ou do .ppam) e lixo sao ignorados. Pura sobre o cache da sessao.
Public Function PptLinks_Remember(lines As String, Optional pres As Object) As Long
    Dim a() As String, f() As String, i As Long, pk As String
    PptLinks_EnsureStates
    If Len(lines) = 0 Or Left$(lines, 1) = "!" Then Exit Function
    pk = PptLinks_PresKey(pres)
    a = Split(lines, vbLf)
    For i = 0 To UBound(a)
        f = Split(a(i), "|")
        If UBound(f) = 5 Then
            If Len(f(0)) > 0 And Len(f(0)) < 10 And Not f(0) Like "*[!0-9]*" Then
                mStates(PptLinks_Key(pk, f(0), f(1))) = f(5)
                PptLinks_Remember = PptLinks_Remember + 1
            End If
        End If
    Next
End Function

' Esquece o estado lembrado das formas dadas (com e sem a apresentacao na chave). Nunca propaga.
Private Sub PptLinks_ForgetShapes(pres As Object, shapes As Collection)
    Dim s As Object, pk As String, k As String
    On Error Resume Next
    PptLinks_EnsureStates
    pk = PptLinks_PresKey(pres)
    For Each s In shapes
        k = PptLinks_Key(pk, s.Parent.SlideID, s.Name)
        If mStates.Exists(k) Then mStates.Remove k
        k = PptLinks_Key("", s.Parent.SlideID, s.Name)
        If mStates.Exists(k) Then mStates.Remove k
    Next
    Err.Clear
End Sub

' Esquece todos os estados lembrados de uma apresentacao (fechada). Nunca propaga.
Private Sub PptLinks_ForgetPres(presKey As String)
    Dim k As Variant, gone As New Collection
    On Error Resume Next
    If mStates Is Nothing Or Len(presKey) = 0 Then Exit Sub
    For Each k In mStates.Keys
        If Left$(CStr(k), Len(presKey) + 1) = presKey & "|" Then gone.Add CStr(k)
    Next
    For Each k In gone
        mStates.Remove CStr(k)
    Next
    Err.Clear
End Sub

' Linhas do dialogo: Array(slideIndex, slideId, nome, tipo, modo, estado, pasta) de cada forma vinculada solta, na ordem dos slides. Estado:
' "new" pela etiqueta (a bandeira); senao o lembrado do core (o da apresentacao, depois o sem apresentacao; "new" lembrado sem a etiqueta =
' "ok": o core marcou e a atualizacao a tirou); senao "". So etiquetas e o vinculo (nunca o arquivo). Nunca propaga.
Public Function PptLinks_Collect(pres As Object) As Collection
    Dim out As New Collection, sl As Object, s As Object, kind As String, mode As String, st As String, pk As String, k As String
    Set PptLinks_Collect = out
    On Error Resume Next
    PptLinks_EnsureStates
    pk = PptLinks_PresKey(pres)
    For Each sl In pres.Slides
        For Each s In sl.Shapes
            If PptTqq_IsLinkedShape(s) Then
                If Len(PptTqq_Tag(s, PPT_TAG_CHART)) > 0 Then
                    kind = "chart"
                ElseIf PptTqq_Tag(s, PPT_TAG_TEXT) = "1" Then
                    kind = "text"
                Else
                    kind = "table"
                End If
                If PptTqq_Tag(s, PPT_TAG_AUTO) = "1" Then mode = "auto" Else mode = "manual"
                st = ""
                If PptTqq_Tag(s, PPT_TAG_STATE) = "new" Then
                    st = "new"
                Else
                    k = PptLinks_Key(pk, sl.SlideID, s.Name)
                    If Not mStates.Exists(k) Then k = PptLinks_Key("", sl.SlideID, s.Name)
                    If mStates.Exists(k) Then
                        st = CStr(mStates(k))
                        If st = "new" Then st = "ok"
                    End If
                End If
                out.Add Array(CLng(sl.SlideIndex), CLng(sl.SlideID), CStr(s.Name), kind, mode, st, PptTqq_SourcePath(PptTqq_LinkSourceOf(s), PptTqq_Tag(s, PPT_TAG_BOOK)))
            End If
            Err.Clear
        Next
    Next
    Err.Clear
End Function

Public Function PptLinks_Filter(rows As Collection, filterIndex As Long) As Collection
    Dim out As New Collection, r As Variant, ok As Boolean
    For Each r In rows
        Select Case filterIndex
            Case 1: ok = (r(5) = "new")
            Case 2: ok = (r(5) = "ok")
            Case 3: ok = (r(4) = "auto")
            Case 4: ok = (r(5) = "missing" Or r(5) = "nosource" Or r(5) = "unlinked")
            Case Else: ok = True
        End Select
        If ok Then out.Add r
    Next
    Set PptLinks_Filter = out
End Function

Public Function PptLinks_HeadText(rows As Collection) As String
    Dim r As Variant, nNew As Long
    For Each r In rows
        If r(5) = "new" Then nNew = nNew + 1
    Next
    PptLinks_HeadText = rows.Count & " elemento(s) vinculado(s) ao Excel; " & nNew & " com dados novos."
End Function

' O que a lista mostra, numa cadeia: o dialogo so redesenha quando ela muda (sem piscar a cada selecao). Pura.
Public Function PptLinks_Signature(rows As Collection) As String
    Dim r As Variant, s As String, i As Long
    For Each r In rows
        For i = 0 To 6
            s = s & CStr(r(i)) & vbTab
        Next
        s = s & vbLf
    Next
    PptLinks_Signature = s
End Function

' Formas das linhas (slide pelo SlideID, forma pelo nome). Linhas que nao acham mais a forma sao ignoradas. Nunca propaga.
Public Function PptLinks_Shapes(pres As Object, rows As Collection) As Collection
    Dim out As New Collection, r As Variant, sl As Object, s As Object
    On Error Resume Next
    For Each r In rows
        Set sl = Nothing: Set s = Nothing
        Set sl = pres.Slides.FindBySlideID(CLng(r(1)))
        If Not sl Is Nothing Then Set s = sl.Shapes(CStr(r(2)))
        If Not s Is Nothing Then out.Add s
        Err.Clear
    Next
    Set PptLinks_Shapes = out
End Function

' Desvincular: tira do indice dos automaticos e apaga as etiquetas TQQ_* das formas (a forma e o vinculo nativo ficam); reconstroi TQQ_LINKIDX.
Public Function PptLinks_Unlink(shapes As Collection) As Long
    Dim s As Object, k As Variant, pres As Object
    If shapes.Count = 0 Then Exit Function
    Set pres = shapes(1).Parent.Parent
    PptTqq_SetAutoShapes shapes, False
    PptLinks_ForgetShapes pres, shapes
    For Each s In shapes
        For Each k In Split(PPTLINKS_TAGS, ",")
            On Error Resume Next
            s.Tags.Delete CStr(k)
            Err.Clear
            On Error GoTo 0
        Next
        PptLinks_Unlink = PptLinks_Unlink + 1
    Next
    PptTqq_RebuildLinkIdx pres
End Function

' Uma acao do dialogo sobre as linhas dadas. Devolve o texto da linha de status (vazio quando a acao ja mostrou a propria mensagem). Nunca propaga.
Public Function PptLinks_Do(pres As Object, what As String, rows As Collection) As String
    Dim shapes As Collection, charts As Collection, s As Object, msg As String, entries As String, n As Long, nNew As Long, nBad As Long, r As Variant
    On Error GoTo Fail
    If what <> "check" Then
        If rows Is Nothing Then PptLinks_Do = PPTLINKS_MSG_SELECT: Exit Function
        If rows.Count = 0 Then PptLinks_Do = PPTLINKS_MSG_SELECT: Exit Function
    End If
    Select Case what
        Case "check"
            msg = PptTqq_LinkStates(pres, "", True, True)
            If Left$(msg, 1) = "!" Then PptLinks_Do = Mid$(msg, 2): Exit Function
            n = PptLinks_Remember(msg, pres)
            For Each r In PptLinks_Collect(pres)
                If r(5) = "new" Then nNew = nNew + 1
                If r(5) = "missing" Or r(5) = "nosource" Or r(5) = "unlinked" Then nBad = nBad + 1
            Next
            PptRx_InvalidateElem
            PptLinks_Do = "Verificados: " & n & ". Com dados novos: " & nNew & ". Com problema: " & nBad & "."
            Exit Function
    End Select
    Set shapes = PptLinks_Shapes(pres, rows)
    If shapes.Count = 0 Then PptLinks_Do = PPTLINKS_MSG_SELECT: Exit Function
    Select Case what
        Case "update"
            For Each s In shapes
                PptTqq_EnsureUniqueNamePublic s
            Next
            entries = PptTqq_EntriesOf(shapes)
            msg = PptTqq_SyncEntriesMsg(pres, entries, True)
            PptLinks_ForgetShapes pres, shapes                      ' o estado velho nao vale mais (sem o core: "Nao verificado")
            PptLinks_Remember PptTqq_LinkStates(pres, entries, False), pres
            If Left$(msg, 1) = "!" Then msg = Mid$(msg, 2)
            PptLinks_Do = Replace(msg, vbLf, " ")
        Case "auto"
            PptTqq_SetAutoShapes shapes, True
            entries = PptTqq_EntriesOf(shapes)
            msg = PptTqq_SyncEntriesMsg(pres, entries, True, True)
            PptLinks_ForgetShapes pres, shapes
            PptLinks_Remember PptTqq_LinkStates(pres, entries, False), pres
            If Left$(msg, 1) = "!" Then msg = Mid$(msg, 2)
            PptLinks_Do = "Automático: " & shapes.Count & " elemento(s). " & Replace(msg, vbLf, " ")
        Case "manual"
            PptTqq_SetAutoShapes shapes, False
            PptLinks_Do = "Manual: " & shapes.Count & " elemento(s)."
        Case "unlink"
            If PptTqq_Ask("Desvincular " & shapes.Count & " elemento(s)? O ThinQQ deixa de atualizá-los e eles saem desta lista. " & _
                          "Um gráfico continua com o vínculo do próprio PowerPoint com a pasta do Excel (Arquivo > Informações > Editar Links para Arquivos).") <> vbYes Then
                If BUILD_CONFIG <> "Test" Then Exit Function
            End If
            PptLinks_Do = "Desvinculados: " & PptLinks_Unlink(shapes) & "."
        Case "edit"
            Set charts = New Collection
            For Each s In shapes
                If Len(PptTqq_Tag(s, PPT_TAG_CHART)) > 0 Then charts.Add s
            Next
            If charts.Count = 0 Then PptLinks_Do = PPTLINKS_MSG_EDIT_CHART: Exit Function
            PptTqq_EditShapes pres, charts
        Case "goto"
            PptLinks_Goto pres, shapes(1)
    End Select
    PptRx_InvalidateElem
    Exit Function
Fail:
    PptLinks_Do = "Não foi possível concluir: " & Err.Description & " (" & Err.Number & ")"
End Function

' Mostra o slide da forma na janela da apresentacao e seleciona a forma. Nunca propaga.
Private Sub PptLinks_Goto(pres As Object, s As Object)
    Dim w As Object
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then Exit Sub
    If pres.Windows.Count = 0 Then Exit Sub
    Set w = pres.Windows(1)
    w.Activate
    w.View.GotoSlide CLng(s.Parent.SlideIndex)
    s.Select
    Err.Clear
End Sub

' Apresentacao da janela ativa (Nothing sem janela). Nunca propaga.
Public Function PptLinks_ActivePres() As Object
    On Error Resume Next
    If Application.Windows.Count = 0 Then Exit Function
    Set PptLinks_ActivePres = Application.ActiveWindow.Presentation
    Err.Clear
End Function

' Chaves (PptLinks_RowKey) dos elementos vinculados selecionados na janela ativa (dicionario; vazio sem selecao). So etiquetas. Nunca propaga.
Public Function PptLinks_SelectedKeys() As Object
    Dim d As Object, s As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set PptLinks_SelectedKeys = d
    On Error Resume Next
    For Each s In PptTqq_SelectedLinked()
        d(CStr(s.Parent.SlideID) & "|" & LCase$(CStr(s.Name))) = True
    Next
    Err.Clear
End Function

' O dialogo carregado (Nothing se nao ha): procura na colecao UserForms, que so lista os carregados (nunca carrega um).
Private Function PptLinks_Form() As Object
    Dim i As Long
    On Error Resume Next
    For i = 0 To VBA.UserForms.Count - 1
        If TypeName(VBA.UserForms(i)) = PPTLINKS_FORMNAME Then Set PptLinks_Form = VBA.UserForms(i): Exit Function
    Next
End Function

' Quantos dialogos Vinculos estao carregados (testes).
Public Function PptLinks_FormCount() As Long
    Dim i As Long
    On Error Resume Next
    For i = 0 To VBA.UserForms.Count - 1
        If TypeName(VBA.UserForms(i)) = PPTLINKS_FORMNAME Then PptLinks_FormCount = PptLinks_FormCount + 1
    Next
End Function

' Abre (ou traz e atualiza) o dialogo para a apresentacao ativa. Os estados vem do Excel JA aberto com o core (nunca cria um Excel nem abre
' pasta fechada; sem ele: "Nao verificado"). Nunca propaga erro.
Public Sub PptLinks_Show()
    Dim p As Object
    On Error GoTo Fail
    Set p = PptLinks_ActivePres()
    If p Is Nothing Then PptTqq_Msg PPT_MSG_NO_PRES: Exit Sub
    If BUILD_CONFIG = "Test" Then Exit Sub
    PptLinks_Remember PptTqq_LinkStates(p, "", False), p
    frmTqqLinks.Prepare p
    mOpen = True
    frmTqqLinks.Show vbModeless
    Exit Sub
Fail:
    PptTqq_Msg "Não foi possível abrir os Vínculos: " & Err.Description & " (" & Err.Number & ")", vbExclamation
End Sub

' O formulario saiu (Fechar, X ou Unload). Chamado pelo proprio formulario.
Public Sub PptLinks_FormClosed()
    mOpen = False
End Sub

' Selecao, slide ou janela mudou no PowerPoint (clsPptEvents): o dialogo aberto acompanha. Sem dialogo aberto, nada (nunca o carrega).
' Nunca chama o Excel nem propaga erro.
Public Sub PptLinks_OnSelChange()
    Dim f As Object
    If Not mOpen Then Exit Sub
    On Error Resume Next
    Set f = PptLinks_Form()
    If f Is Nothing Then mOpen = False: Exit Sub
    f.SyncView
    Err.Clear
End Sub

' Apresentacao fechada: os estados dela saem do cache e o dialogo dela e descarregado (nao fica apontando para um objeto morto). Nunca propaga.
Public Sub PptLinks_OnClose(pres As Object)
    Dim i As Long, f As Object, k As String
    On Error Resume Next
    k = PptLinks_PresKey(pres)
    PptLinks_ForgetPres k
    For i = VBA.UserForms.Count - 1 To 0 Step -1
        Set f = Nothing
        Set f = VBA.UserForms(i)
        If TypeName(f) = PPTLINKS_FORMNAME Then
            If CStr(f.PresKey) = k Then Unload f
        End If
    Next
    Err.Clear
End Sub

' PowerPoint (ou o suplemento) fechando: descarrega todo dialogo Vinculos. Nunca propaga.
Public Sub PptLinks_CloseAll()
    Dim i As Long, f As Object
    On Error Resume Next
    For i = VBA.UserForms.Count - 1 To 0 Step -1
        Set f = Nothing
        Set f = VBA.UserForms(i)
        If TypeName(f) = PPTLINKS_FORMNAME Then Unload f
    Next
    mOpen = False
    Err.Clear
End Sub
