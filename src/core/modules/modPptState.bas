Attribute VB_Name = "modPptState"
Option Explicit
' Estado dos elementos do ThinQQ no PowerPoint (1.11.0; spec 2026-10-23 secao 8.1). Como no think-cell, o vinculo e MANUAL por padrao e o slide
' mostra quando a pasta tem dados novos (a bandeira). O core compara, SEM ALTERAR NADA no Excel (o Desfazer do usuario fica: nada de
' redimensionar, de Core_SafeRefresh nem de salvar), o que o slide mostra (TQQ_FP do grafico, TQQ_TLAST do texto, TQQ_TBLAST da tabela) com o
' que a pasta aberta daria agora, e grava na forma TQQ_STATE = "new" ou a apaga. TQQ_LINKIDX (etiqueta da apresentacao, formato PptIdx_*)
' lista TODOS os elementos vinculados; TQQ_AUTOIDX continua so com os automaticos. Gravar estas etiquetas preserva Presentation.Saved (sao
' estado da sessao, nao alteracao do usuario).
Public Const PPTSTATE_TAG As String = "TQQ_STATE"
Public Const PPTSTATE_TAG_LINKIDX As String = "TQQ_LINKIDX"
Public Const PPTSTATE_NEW As String = "new"              ' dados novos na pasta (a bandeira)
Public Const PPTSTATE_OK As String = "ok"                ' o slide mostra o que a pasta tem
Public Const PPTSTATE_UNKNOWN As String = "unknown"      ' sem referencia gravada (TQQ_FP ausente) ou pasta fechada nao verificada
Public Const PPTSTATE_MISSING As String = "missing"      ' pasta de origem nao encontrada
Public Const PPTSTATE_NOSOURCE As String = "nosource"    ' grafico, celula ou intervalo nao encontrado na pasta
Public Const PPTSTATE_UNLINKED As String = "unlinked"    ' grafico sem o vinculo nativo
Private Const PPTSTATE_MAX_FIELD As Long = 1024         ' PptState_Lines: teto do nome e da pasta em cada linha (texto vindo do deck)
Private mTestMarks As Long                             ' gancho de teste: etiquetas TQQ_STATE gravadas/apagadas por PptState_SetTag
Private mTestProbes As Long                            ' gancho de teste: pastas fechadas sondadas no disco por PptState_Lines (so com Verificar)

Public Function PptState_MarksForTest() As Long
    PptState_MarksForTest = mTestMarks
End Function

Public Sub PptState_ResetTestHooks()
    mTestMarks = 0
    mTestProbes = 0
End Sub

' Etiqueta lida sem engolir o PowerPoint ocupado (sobe); outro erro ou etiqueta ausente = "".
Private Function PptState_TagStrict(o As Object, key As String) As String
    Dim v As Variant, en As Long, ed As String
    On Error Resume Next
    v = o.Tags.Item(key)
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    If Ppt_IsBusyError(en) Then Err.Raise en, "PptState_TagStrict", ed
    If en = 0 And Not IsNull(v) And Not IsObject(v) Then PptState_TagStrict = CStr(v)
End Function

' Tipo do elemento (a regra de PptSync_ItemKind) sem engolir o ocupado.
Private Function PptState_KindStrict(shp As Object) As String
    If Len(PptState_TagStrict(shp, PPTSYNC_TAG_CHART)) > 0 Then PptState_KindStrict = "chart": Exit Function
    If PptState_TagStrict(shp, PPTTEXT_TAG) = "1" Then PptState_KindStrict = "text": Exit Function
    If PptState_TagStrict(shp, PPTTABLE_TAG) = "1" Then PptState_KindStrict = "table"
End Function

' Intervalo de um texto/tabela vinculado na pasta wb com as MESMAS regras de seguranca de PptText_Sync/PptTable_Sync (revisao final E, I1): so
' um nome com o prefixo dado (criado pela acao de vincular) e seguido; o endereco de reserva so numa pasta com a marca PPTLINK_MARK. Nothing se
' nao acha. PowerPoint ocupado sobe.
Private Function PptState_LinkRange(shp As Object, wb As Workbook, prefix As String) As Range
    Dim nm As String, ref As String
    nm = PptState_TagStrict(shp, PPTLINK_TAG_NAME)
    If Left$(nm, Len(prefix)) <> prefix Then nm = ""
    If PptLink_HasMark(wb) Then ref = PptState_TagStrict(shp, PPTLINK_TAG_REF)
    If Len(nm) = 0 And Len(ref) = 0 Then Exit Function
    On Error Resume Next
    Set PptState_LinkRange = PptLink_ResolveRange(wb, nm, ref)
    Err.Clear
End Function

' Estado do elemento pela pasta wb (aberta). So LE: nada e gravado no Excel nem no slide. PowerPoint ocupado sobe; outro erro = UNKNOWN (registrado).
Public Function PptState_Of(shp As Object, wb As Workbook) As String
    Dim co As ChartObject, fp As String, rng As Range, v As String, cut As Boolean, oldM As Variant, newM As Variant, en As Long, ed As String
    On Error GoTo Fail
    Select Case PptState_KindStrict(shp)
        Case "chart"
            Set co = Registry_FindChartObject(wb, PptState_TagStrict(shp, PPTSYNC_TAG_CHART))
            If co Is Nothing Then PptState_Of = PPTSTATE_NOSOURCE: Exit Function
            fp = PptState_TagStrict(shp, PPTSYNC_TAG_FP)
            If Len(fp) = 0 Then PptState_Of = PPTSTATE_UNKNOWN: Exit Function
            If PptSync_Fingerprint(co, shp) = fp Then PptState_Of = PPTSTATE_OK Else PptState_Of = PPTSTATE_NEW
        Case "text"
            Set rng = PptState_LinkRange(shp, wb, "TQQ_X_")
            If rng Is Nothing Then PptState_Of = PPTSTATE_NOSOURCE: Exit Function
            v = PptText_SafeValue(PptText_CellText(rng), cut)
            If v = PptText_TagDec(PptState_TagStrict(shp, PPTTEXT_TAG_LAST)) Then PptState_Of = PPTSTATE_OK Else PptState_Of = PPTSTATE_NEW
        Case "table"
            Set rng = PptState_LinkRange(shp, wb, PPTTABLE_NAME_PREFIX)
            If rng Is Nothing Then PptState_Of = PPTSTATE_NOSOURCE: Exit Function
            ' a reserva TQQ_TREF vem da apresentacao: um intervalo fora dos limites da tabela (75 x 25, sem mescladas, uma area) nunca e lido
            ' celula a celula (PptTable_Sync recusa igual); sem estado
            If Len(PptTable_RefusalText(rng)) > 0 Then PptState_Of = PPTSTATE_UNKNOWN: Exit Function
            oldM = PptTable_Parse(PptState_TagStrict(shp, PPTTABLE_TAG_LAST))
            If IsEmpty(oldM) Then PptState_Of = PPTSTATE_UNKNOWN: Exit Function
            newM = PptTable_TextMatrix(rng)
            If UBound(oldM, 1) <> UBound(newM, 1) Or UBound(oldM, 2) <> UBound(newM, 2) Then PptState_Of = PPTSTATE_NEW: Exit Function
            If PptTable_Diff(oldM, newM).Count = 0 Then PptState_Of = PPTSTATE_OK Else PptState_Of = PPTSTATE_NEW
        Case Else
            PptState_Of = PPTSTATE_UNKNOWN
    End Select
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error GoTo 0                                     ' sem isto o GoTo -1 rearma o Fail e o Raise voltaria para ca (laco)
    If Ppt_IsBusyError(en) Then Err.Raise en, "PptState_Of", ed
    Errors_Log "PptState_Of" & vbTab & en & vbTab & ed
    PptState_Of = PPTSTATE_UNKNOWN
End Function

' Grava (isNew) ou apaga TQQ_STATE so quando muda (um valor diferente de "new", vindo de um deck alheio, nao e bandeira e fica). True se gravou
' ou apagou. PowerPoint ocupado sobe; outro erro e registrado (False).
Public Function PptState_SetTag(shp As Object, isNew As Boolean) As Boolean
    Dim cur As String, en As Long, ed As String
    cur = PptState_TagStrict(shp, PPTSTATE_TAG)
    If isNew = (cur = PPTSTATE_NEW) Then Exit Function
    If Not isNew Then
        If cur <> PPTSTATE_NEW Then Exit Function
    End If
    On Error Resume Next
    If isNew Then shp.Tags.Add PPTSTATE_TAG, PPTSTATE_NEW Else shp.Tags.Delete PPTSTATE_TAG
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    If Ppt_IsBusyError(en) Then Err.Raise en, "PptState_SetTag", ed
    If en <> 0 Then Errors_Log "PptState_SetTag" & vbTab & en & vbTab & ed: Exit Function
    PptState_SetTag = True
    mTestMarks = mTestMarks + 1
End Function

' Presentation.Saved antes de gravar etiquetas de estado (False se ilegivel). Nunca propaga erro.
Public Function PptState_SavedOf(pres As Object) As Boolean
    On Error Resume Next
    PptState_SavedOf = (CLng(pres.Saved) <> 0)
    Err.Clear
End Function

' Devolve Saved = True se a apresentacao estava salva antes (P-U1). Nunca propaga erro.
Public Sub PptState_RestoreSaved(pres As Object, wasSaved As Boolean)
    On Error Resume Next
    If wasSaved Then
        If CLng(pres.Saved) = 0 Then pres.Saved = -1
    End If
    Err.Clear
End Sub

' Acrescenta a forma (elemento do ThinQQ recem-criado, colado ou vinculado) ao TQQ_LINKIDX da apresentacao, com o nome unico no slide (o indice
' acha a forma pelo nome). Preserva Saved. Nunca propaga erro (falha registrada: a sincronizacao completa e a abertura no .ppam reconstroem).
Public Sub PptLinkIdx_Add(shp As Object)
    Dim sl As Object, pres As Object, idx As String, was As Boolean
    On Error GoTo Fail
    Set sl = shp.Parent: Set pres = sl.Parent
    was = PptState_SavedOf(pres)
    PptSync_EnsureUniqueName shp
    idx = PptIdx_Set(PptSync_Tag(pres, PPTSTATE_TAG_LINKIDX), CLng(sl.SlideID), CStr(shp.Name), PptSync_ItemBook(shp), True)
    pres.Tags.Add PPTSTATE_TAG_LINKIDX, idx
    PptState_RestoreSaved pres, was
    Exit Sub
Fail:
    Errors_Log "PptLinkIdx_Add" & vbTab & Err.Number & vbTab & Err.Description
    If Not pres Is Nothing Then PptState_RestoreSaved pres, was        ' revisao final M11
End Sub

' Reconstroi TQQ_LINKIDX com todos os itens dados (todas as formas do ThinQQ da apresentacao); so grava se mudou; preserva Saved. Cada forma
' isolada (a que falha fica fora e e registrada). Nunca propaga erro.
Public Sub PptLinkIdx_Rebuild(pres As Object, items As Collection)
    Dim shp As Object, out As New Collection, s As String, was As Boolean
    On Error GoTo Done
    was = PptState_SavedOf(pres)                    ' antes de renomear (revisao final M7): a renomeacao de duplicadas tambem nao suja
    For Each shp In items
        On Error Resume Next
        PptSync_EnsureUniqueName shp
        out.Add Array(CLng(shp.Parent.SlideID), CStr(shp.Name), PptSync_ItemBook(shp))
        If Err.Number <> 0 Then Errors_Log "PptLinkIdx_Rebuild: forma fora do indice" & vbTab & Err.Number & vbTab & Err.Description
        Err.Clear
        On Error GoTo Done
    Next
    s = PptIdx_Join(out)
    If s <> PptSync_Tag(pres, PPTSTATE_TAG_LINKIDX) Then
        If Len(s) = 0 Then PptSync_TagDelete pres, PPTSTATE_TAG_LINKIDX Else pres.Tags.Add PPTSTATE_TAG_LINKIDX, s
    End If
    PptState_RestoreSaved pres, was
    Exit Sub
Done:
    Errors_Log "PptLinkIdx_Rebuild" & vbTab & Err.Number & vbTab & Err.Description
    On Error Resume Next
    PptState_RestoreSaved pres, was
End Sub

' Elementos MANUAIS (sem TQQ_AUTO = "1") de pres cuja pasta de origem esta em dirty, pelo TQQ_LINKIDX (nunca varre as formas; sem o indice:
' vazia). A pasta da entrada pre-filtra antes de qualquer chamada ao PowerPoint; a pasta REAL da forma tambem tem de estar em dirty. Entradas
' velhas sao ignoradas (a sincronizacao completa e a abertura reconstroem o indice). PowerPoint ocupado sobe.
Public Function PptState_ManualFor(pres As Object, dirty As Object) As Collection
    Dim out As New Collection, e As Variant, sl As Object, shp As Object, en As Long, seen As Object, k As String, idx As String
    Set PptState_ManualFor = out
    idx = PptState_TagStrict(pres, PPTSTATE_TAG_LINKIDX)
    If Len(idx) = 0 Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For Each e In PptIdx_Parse(idx)
        If dirty.Exists(CStr(e(2))) Then
            Set sl = Nothing: Set shp = Nothing
            On Error Resume Next
            Set sl = pres.Slides.FindBySlideID(CLng(e(0)))
            If Not sl Is Nothing Then Set shp = sl.Shapes(CStr(e(1)))
            en = Err.Number
            Err.Clear
            On Error GoTo 0
            If Ppt_IsBusyError(en) Then Err.Raise en, "PptState_ManualFor", "PowerPoint ocupado"
            If Not shp Is Nothing Then
                If PptState_TagStrict(shp, PPTSYNC_TAG_AUTO) <> "1" Then
                    If Len(PptState_KindStrict(shp)) > 0 Then
                        k = CStr(e(0)) & "|" & LCase$(CStr(e(1)))
                        If Not seen.Exists(k) Then
                            seen.Add k, True
                            If dirty.Exists(PptSync_ItemBook(shp)) Then out.Add shp
                        End If
                    End If
                End If
            End If
        End If
    Next
End Function

' Caminho remoto (a mesma regra de PptTqq_IsRemotePath do .ppam): rede (\\servidor\..., \\?\UNC\, WebDAV, barras normais ou mistas),
' http(s) ou file:. \\?\C:\ e letras de unidade sao locais (unidade de rede mapeada nao e distinguivel: limite documentado). Pura.
Public Function PptState_IsRemotePath(p As String) As Boolean
    Dim t As String
    t = Replace(LCase$(Trim$(p)), "/", "\")
    If Left$(t, 5) = "file:" Or Left$(t, 7) = "http:\\" Or Left$(t, 8) = "https:\\" Then PptState_IsRemotePath = True: Exit Function
    If Left$(t, 2) <> "\\" Then Exit Function
    If Left$(t, 8) = "\\?\unc\" Then PptState_IsRemotePath = True: Exit Function
    If (Left$(t, 4) = "\\?\" Or Left$(t, 4) = "\\.\") And Mid$(t, 6, 1) = ":" Then Exit Function
    PptState_IsRemotePath = True
End Function

' Somente leitura (True tambem se ilegivel: na duvida nao grava). Nunca propaga erro.
Private Function PptState_IsReadOnly(pres As Object) As Boolean
    PptState_IsReadOnly = True
    On Error Resume Next
    PptState_IsReadOnly = (CLng(pres.ReadOnly) <> 0)
    Err.Clear
End Function

Public Function PptState_ProbesForTest() As Long
    PptState_ProbesForTest = mTestProbes
End Function

Private Function PptState_PresKey(pres As Object) As String
    On Error Resume Next
    PptState_PresKey = LCase$(CStr(pres.FullName))
    If Len(PptState_PresKey) = 0 Then PptState_PresKey = "?" & CStr(ObjPtr(pres))
    Err.Clear
End Function

Private Sub PptState_RestoreAll(was As Object)
    Dim k As Variant, a As Variant, p As Object
    On Error Resume Next
    For Each k In was.Keys
        a = was(k)
        Set p = a(0)                                    ' Variant nao passa ByRef para "As Object" (erro de compilacao)
        PptState_RestoreSaved p, CBool(a(1))
    Next
    Err.Clear
End Sub

' Marca (TQQ_STATE) os elementos dados pela pasta ABERTA neste Excel (nunca abre pasta): "new" se o slide esta atras, apagada se esta em dia;
' estado desconhecido ou sem origem nao mexe na etiqueta. Saved de cada apresentacao preservado. Devolve quantos estao com dados novos.
' PowerPoint ocupado (ou falha estrutural) sobe depois de devolver o Saved.
Public Function PptState_MarkItems(items As Collection) As Long
    Dim shp As Object, wb As Workbook, st As String, pres As Object, was As Object, k As String, en As Long, ed As String
    Set was = CreateObject("Scripting.Dictionary")
    On Error GoTo Fail
    For Each shp In items
        Set wb = PptSync_FindOpenBook(PptSync_ItemBook(shp))
        If Not wb Is Nothing Then
            st = PptState_Of(shp, wb)
            If st = PPTSTATE_NEW Or st = PPTSTATE_OK Then
                Set pres = shp.Parent.Parent
                k = PptState_PresKey(pres)
                If Not was.Exists(k) Then was.Add k, Array(pres, PptState_SavedOf(pres))
                PptState_SetTag shp, (st = PPTSTATE_NEW)
                If st = PPTSTATE_NEW Then PptState_MarkItems = PptState_MarkItems + 1
            End If
        End If
    Next
    PptState_RestoreAll was
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptState_RestoreAll was
    On Error GoTo 0
    Err.Raise en, "PptState_MarkItems", ed
End Function

' Estado de cada elemento (entries no formato do indice; "" = todos os elementos da apresentacao, por varredura, e o TQQ_LINKIDX e
' reconstruido). Linha: SlideID|nome|pasta|tipo|modo|estado (campos sem os separadores, PptIdx_SafeField; linhas por vbLf); tipo = "chart:" +
' TQQ_KIND, "text" ou "table"; modo = auto|manual; estado = PPTSTATE_*. Pasta fechada: com verifyClosed, aberta SOMENTE LEITURA com as macros
' desligadas (PptSync_OpenBook) e fechada no fim ("unknown" se existe e nao abriu, "missing" se nao existe); sem, "unknown" SEM tocar o disco
' (revisao final M3: como na 1.10.0, so "Verificar" sonda pastas fechadas; uma unidade mapeada fora da VPN travaria a abertura). Atualiza
' TQQ_STATE e, com entries = "", o TQQ_LINKIDX (Saved preservado); apresentacao somente leitura: nada e gravado nem renomeado, os estados so
' sao calculados (revisao final M2). PowerPoint ocupado sobe.
Public Function PptState_Lines(pres As Object, entries As String, verifyClosed As Boolean) As String
    Dim items As Collection, sl As Object, shp As Object, opened As New Collection, books As Object, out() As String, n As Long
    Dim path As String, wb As Workbook, st As String, kind As String, mode As String, was As Boolean, en As Long, ed As String, ro As Boolean
    On Error GoTo Fail
    was = PptState_SavedOf(pres)
    ro = PptState_IsReadOnly(pres)
    If Len(entries) = 0 Then
        Set items = New Collection
        For Each sl In pres.Slides
            For Each shp In sl.Shapes
                If Len(PptSync_ItemKind(shp)) > 0 Then items.Add shp
            Next
        Next
        If Not ro Then PptLinkIdx_Rebuild pres, items
    Else
        Set items = PptSync_ResolveEntries(pres, entries)
    End If
    If items.Count > 0 Then
        Set books = CreateObject("Scripting.Dictionary"): books.CompareMode = 1
        ReDim out(1 To items.Count)
        For Each shp In items
            kind = PptSync_ItemKind(shp)
            If PptSync_Tag(shp, PPTSYNC_TAG_AUTO) = "1" Then mode = "auto" Else mode = "manual"
            path = PptSync_ItemBook(shp)
            st = ""
            If kind = "chart" Then
                If Ppt_LinkState(shp) = 0 Then st = PPTSTATE_UNLINKED
            End If
            If Len(st) = 0 Then
                If Not books.Exists(path) Then
                    Set wb = PptSync_FindOpenBook(path)
                    If wb Is Nothing And verifyClosed And Len(path) > 0 Then
                        If BUILD_CONFIG = "Test" Then mTestProbes = mTestProbes + 1
                        If PptSync_FileExists(path) Then Set wb = PptSync_OpenBook(path, opened)
                    End If
                    books.Add path, wb
                End If
                Set wb = books(path)
                If wb Is Nothing Then
                    If Len(path) = 0 Then
                        st = PPTSTATE_MISSING
                    ElseIf Not verifyClosed Then
                        st = PPTSTATE_UNKNOWN                   ' sem "Verificar": o disco nunca e tocado (rede, web, unidade mapeada)
                    ElseIf PptSync_FileExists(path) Then
                        st = PPTSTATE_UNKNOWN
                    Else
                        st = PPTSTATE_MISSING
                    End If
                Else
                    st = PptState_Of(shp, wb)
                End If
            End If
            If st = PPTSTATE_NEW Or st = PPTSTATE_OK Then
                If Not ro Then PptState_SetTag shp, (st = PPTSTATE_NEW)
            End If
            If kind = "chart" Then kind = "chart:" & PptSync_OptValue(PptSync_Tag(shp, PPTSYNC_TAG_KIND))
            n = n + 1
            ' campos vindos da apresentacao (de quem a fez): sem separadores e com teto, para a lista do .ppam nunca receber um texto gigante
            out(n) = CStr(shp.Parent.SlideID) & "|" & Left$(PptIdx_SafeField(CStr(shp.Name)), PPTSTATE_MAX_FIELD) & "|" & Left$(PptIdx_SafeField(path), PPTSTATE_MAX_FIELD) & "|" & kind & "|" & mode & "|" & st
        Next
        PptState_Lines = Join(out, vbLf)
    End If
    PptSync_CloseOpened opened
    PptState_RestoreSaved pres, was
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptSync_CloseOpened opened
    PptState_RestoreSaved pres, was
    On Error GoTo 0
    Err.Raise en, "PptState_Lines", ed
End Function
