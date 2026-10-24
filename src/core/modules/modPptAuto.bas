Attribute VB_Name = "modPptAuto"
Option Explicit
' Atualizacao automatica dos elementos do ThinQQ no PowerPoint (spec 2026-10-15 secao 5.1). Opcoes em
' HKCU\Software\VB and VBA Program Settings\ThinQQ\PowerPoint: AutoUpdate (padrao "1": o motor age nas formas marcadas TQQ_AUTO) e AutoNew
' (padrao "0": o vinculo nasce manual, como no think-cell). No build Test so memoria (nunca o registro).
Private Const PPTAUTO_SECTION As String = "PowerPoint"
' Motor (1.6.0): eventos do Excel (clsAppEvents: SheetChange, SheetCalculate, WorkbookAfterSave) marcam a pasta e (re)agendam um OnTime
' para PPTAUTO_DEBOUNCE_S depois da ULTIMA mudanca; o disparo (PptAuto_Tick) sincroniza so as formas automaticas daquela pasta nas
' apresentacoes do PowerPoint JA aberto. Um unico OnTime vivo (cancelado sempre pelo MESMO horario; nada empilha) e Core_OnUnload o cancela
' (o OnTime reabriria o core descarregado). No build Test o OnTime real so existe quando o teste pede (PptAuto_SetNoTimerForTest False).
Public Const PPTAUTO_DEBOUNCE_S As Long = 2                 ' espera depois da ULTIMA mudanca (cada evento reagenda)
Public Const PPTAUTO_MAX_RETRIES As Long = 5                ' PowerPoint ocupado: novas tentativas antes de desistir
Public Const PPTAUTO_MAX_WAIT_S As Long = 10                ' pasta que nunca para de mudar (RTD): o OnTime vivo deixa de ser adiado
Private Const PPTAUTO_ERR_BUSY As Long = -2147418111        ' 0x80010001 (RPC_E_CALL_REJECTED): adiar como PowerPoint ocupado
Private Const PPTAUTO_HALF_S As Double = 0.5 / 86400        ' meio segundo em Date: horarios iguais "no segundo"
Private Const PPTAUTO_PROGRESS As String = "ThinQQ: atualizando gráfico "     ' inicio do texto de progresso de PptSync_Items
Private Const PPTAUTO_SUM_NONE As String = "Gráficos do ThinQQ atualizados: 0."  ' inicio do resumo sem nada atualizado
Private mTestEnabled As Variant                 ' gancho: Empty = padrao
Private mTestNewAuto As Variant
Private mDirtySince As Date                                 ' horario da primeira marca desde que a fila esvaziou (0 = fila vazia)
Private mSbSaved As Boolean                                 ' o ciclo guardou a barra de status (mSbPrev) antes de sincronizar
Private mSbPrev As Variant
Private mLogCount As Long                                   ' linhas registradas pelo motor (lido so pelos testes)
Private mTestNow As Date                                    ' gancho: relogio do motor (0 = Now)
Private mDue As Date                                        ' horario do OnTime agendado (0 = nenhum): o cancelamento exige o mesmo valor
Private mDueReal As Boolean                                 ' mDue tem um Application.OnTime de verdade (senao so o horario, gancho de teste)
Private mDirty As Object                                    ' Scripting.Dictionary (texto): FullName das pastas que mudaram
Private mInTick As Boolean                                  ' reentrada e eventos do proprio ciclo
Private mRetries As Long
Private mLastResult As String
Private mTimerFailures As Long                              ' OnTime que nao pode ser agendado/cancelado (testes: deve ficar 0)
Private mTestRealTimer As Boolean                           ' gancho: build Test so agenda OnTime real quando o teste pede
Private mTestBusy As Long                                   ' gancho: quantos ciclos simulam PowerPoint ocupado
Private mTestNotify As String                               ' gancho: o proximo ciclo chama Core_Notify/Core_Confirm com este texto (uso unico)
Private mLastMarked As Long                                 ' elementos manuais com bandeira no ultimo ciclo (lido pelos testes)

Public Sub PptAuto_SetFlagsForTest(enabled As Variant, newAuto As Variant)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestEnabled = enabled
    mTestNewAuto = newAuto
End Sub

Public Sub PptAuto_ResetTestHooks()
    mTestEnabled = Empty
    mTestNewAuto = Empty
    mTestBusy = 0
    mTestNotify = ""
    mTimerFailures = 0
    mLogCount = 0
    mTestNow = 0
End Sub

Public Function PptAuto_Enabled() As Boolean
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then
        If IsEmpty(mTestEnabled) Then PptAuto_Enabled = True Else PptAuto_Enabled = CBool(mTestEnabled)
        Exit Function
    End If
    PptAuto_Enabled = (GetSetting(APP_NAME, PPTAUTO_SECTION, "AutoUpdate", "1") <> "0")
End Function

Public Sub PptAuto_SetEnabled(v As Boolean)
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then mTestEnabled = v: Exit Sub
    SaveSetting APP_NAME, PPTAUTO_SECTION, "AutoUpdate", IIf(v, "1", "0")
End Sub

Public Function PptAuto_NewLinksAuto() As Boolean
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then
        If IsEmpty(mTestNewAuto) Then PptAuto_NewLinksAuto = False Else PptAuto_NewLinksAuto = CBool(mTestNewAuto)
        Exit Function
    End If
    PptAuto_NewLinksAuto = (GetSetting(APP_NAME, PPTAUTO_SECTION, "AutoNew", "0") = "1")
End Function

Public Sub PptAuto_SetNewLinksAuto(v As Boolean)
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then mTestNewAuto = v: Exit Sub
    SaveSetting APP_NAME, PPTAUTO_SECTION, "AutoNew", IIf(v, "1", "0")
End Sub

' ---- Motor (Tarefa C2) ----

' Ganchos de teste (inertes fora do build Test; as leituras sao inofensivas).
Public Sub PptAuto_SetNoTimerForTest(v As Boolean)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestRealTimer = Not v
End Sub

Public Sub PptAuto_SetBusyForTest(n As Long)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestBusy = n
End Sub

Public Sub PptAuto_SetInTickForTest(v As Boolean)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mInTick = v
End Sub

' Gancho: o proximo ciclo (depois de marcar "em andamento") chama Core_Notify e Core_Confirm com msg, como um aviso do Core_SafeRefresh
' (tabela de origem mudou, Mesma escala cortada) chamado de dentro da sincronizacao. Uso unico.
Public Sub PptAuto_SetNotifyInTickForTest(msg As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestNotify = msg
End Sub

' Ciclo automatico em andamento (OnTime): Core_Notify e Core_Confirm nao abrem caixa nenhuma (so registram), para o ciclo nunca parar o
' Excel do usuario com uma janela modal que ele nao pediu.
Public Function PptAuto_InTick() As Boolean
    PptAuto_InTick = mInTick
End Function

' Relogio do motor (0 = o relogio real). So no build Test.
Public Sub PptAuto_SetNowForTest(d As Date)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestNow = d
End Sub

' Cancela de verdade um OnTime do ciclo no horario at; True = ele existia. Se era o vivo, o motor o esquece. Nao conta como falha. So no
' build Test (prova que nada empilhou).
Public Function PptAuto_CancelAtForTest(at As Date) As Boolean
    If BUILD_CONFIG <> "Test" Then Exit Function
    On Error Resume Next
    Application.OnTime at, PptAuto_Proc(), , False
    PptAuto_CancelAtForTest = (Err.Number = 0)
    Err.Clear
    If PptAuto_CancelAtForTest And Abs(at - mDue) < PPTAUTO_HALF_S Then mDue = 0: mDueReal = False
End Function

Public Function PptAuto_PendingCountForTest() As Long
    If mDirty Is Nothing Then Exit Function
    PptAuto_PendingCountForTest = mDirty.Count
End Function

Public Function PptAuto_IsScheduledForTest() As Boolean
    PptAuto_IsScheduledForTest = (mDue <> 0)
End Function

Public Function PptAuto_DueForTest() As Date
    PptAuto_DueForTest = mDue
End Function

Public Function PptAuto_LastResultForTest() As String
    PptAuto_LastResultForTest = mLastResult
End Function

Public Function PptAuto_RetriesForTest() As Long
    PptAuto_RetriesForTest = mRetries
End Function

Public Function PptAuto_TimerFailuresForTest() As Long
    PptAuto_TimerFailuresForTest = mTimerFailures
End Function

Public Function PptAuto_LogCountForTest() As Long
    PptAuto_LogCountForTest = mLogCount
End Function

Private Function PptAuto_Now() As Date
    PptAuto_Now = Now
    If BUILD_CONFIG = "Test" Then
        If mTestNow <> 0 Then PptAuto_Now = mTestNow
    End If
End Function

Private Sub PptAuto_Log(msg As String)
    mLogCount = mLogCount + 1
    Errors_Log msg
End Sub

' Pasta que pode ter elementos no PowerPoint: salva (o vinculo aponta para o arquivo) e com a planilha _ThinQQ (graficos do ThinQQ) ou, desde a
' 1.9.0, com a marca PPTLINK_MARK (texto vinculado / tabela: criada junto com os nomes TQQ_X_/TQQ_T_ e salva com eles). Chamada a cada
' SheetChange e SheetCalculate: so buscas diretas, NUNCA um laco pelos nomes da pasta (revisao final E, I3). Nunca propaga erro.
Public Function PptAuto_IsCandidate(wb As Workbook) As Boolean
    Dim ws As Object
    On Error Resume Next
    If wb Is Nothing Then Exit Function
    If Len(wb.Path) = 0 Then Exit Function
    Set ws = wb.Worksheets(CALC_SHEET)
    PptAuto_IsCandidate = Not ws Is Nothing
    If Not PptAuto_IsCandidate Then PptAuto_IsCandidate = PptLink_HasMark(wb)
    Err.Clear
End Function

' Eventos do Excel (SheetChange, SheetCalculate, WorkbookAfterSave): marca a pasta e (re)agenda o ciclo para daqui a PPTAUTO_DEBOUNCE_S.
' Barato quando nada se aplica (pasta sem arquivo ou sem ThinQQ sai antes de ler a opcao). Ignora os eventos do proprio ciclo. Pasta que
' nunca para de mudar: passados PPTAUTO_MAX_WAIT_S desde a primeira marca, o OnTime vivo nao e mais adiado (o ciclo acontece). Nunca
' propaga erro.
Public Sub PptAuto_OnChange(wb As Workbook)
    Dim t As Date
    On Error GoTo Done
    If mInTick Then Exit Sub
    If Not PptAuto_IsCandidate(wb) Then Exit Sub
    If Not PptAuto_Enabled() Then Exit Sub
    If mDirty Is Nothing Then
        Set mDirty = CreateObject("Scripting.Dictionary")
        mDirty.CompareMode = 1
    End If
    t = PptAuto_Now()
    If mDirty.Count = 0 Or mDirtySince = 0 Then mDirtySince = t
    mDirty(wb.FullName) = True
    If mDue <> 0 And DateDiff("s", mDirtySince, t) >= PPTAUTO_MAX_WAIT_S Then Exit Sub
    mRetries = 0
    PptAuto_Schedule PPTAUTO_DEBOUNCE_S
Done:
End Sub

Private Function PptAuto_Proc() As String
    PptAuto_Proc = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PptAuto_Tick"
End Function

' Agenda o ciclo para daqui a secs, cancelando antes o OnTime vivo (um so de cada vez). Now tem resolucao de 1 s: a espera real fica entre
' secs - 1 e secs segundos, e um novo horario igual ao vivo (mesmo segundo) nao cancela nem agenda nada. Sem OnTime possivel: registra e
' fica sem agenda (a proxima mudanca tenta de novo). Nunca propaga erro.
Private Sub PptAuto_Schedule(secs As Long)
    Dim real_ As Boolean, due As Date
    due = PptAuto_Now() + TimeSerial(0, 0, secs)
    If mDue <> 0 And Abs(due - mDue) < PPTAUTO_HALF_S Then Exit Sub
    PptAuto_CancelTimer
    mDue = due
    real_ = True
    If BUILD_CONFIG = "Test" Then real_ = mTestRealTimer
    If Not real_ Then Exit Sub
    On Error Resume Next
    Application.OnTime mDue, PptAuto_Proc()
    If Err.Number <> 0 Then
        mTimerFailures = mTimerFailures + 1
        PptAuto_Log "PptAuto_Schedule" & vbTab & Err.Number & vbTab & Err.Description
        mDue = 0
    Else
        mDueReal = True
    End If
    Err.Clear
End Sub

' Cancela o OnTime vivo pelo MESMO horario e procedimento com que foi agendado. Nunca propaga erro.
Private Sub PptAuto_CancelTimer()
    If mDue = 0 Then Exit Sub
    If mDueReal Then
        On Error Resume Next
        Application.OnTime mDue, PptAuto_Proc(), , False
        If Err.Number <> 0 Then mTimerFailures = mTimerFailures + 1: PptAuto_Log "PptAuto_CancelTimer" & vbTab & Err.Number & vbTab & Err.Description
        Err.Clear
    End If
    mDue = 0
    mDueReal = False
End Sub

' Fila vazia (o ciclo terminou ou desistiu).
Private Sub PptAuto_ClearQueue()
    If Not mDirty Is Nothing Then mDirty.RemoveAll
    mRetries = 0
    mDirtySince = 0
End Sub

' Desliga tudo (Core_OnUnload, testes): desagenda e esvazia a fila. Nunca propaga erro.
Public Sub PptAuto_Cancel()
    On Error Resume Next
    PptAuto_CancelTimer
    Set mDirty = Nothing
    mRetries = 0
    mDirtySince = 0
    mInTick = False
    mSbSaved = False: mSbPrev = Empty
    mTestBusy = 0
    mTestNotify = ""
    mTestRealTimer = False
    mTestNow = 0
    mLastResult = ""
    mLastMarked = 0
End Sub

' Texto em edicao no PowerPoint (o ciclo adia em vez de disputar com o usuario). Nunca propaga erro.
Private Function PptAuto_PptEditingText(app As Object) As Boolean
    On Error Resume Next
    PptAuto_PptEditingText = (CLng(app.ActiveWindow.Selection.Type) = 3)     ' ppSelectionText
    Err.Clear
End Function

' Tira da fila as pastas que nao estao mais abertas neste Excel: o ciclo nunca abre pasta (a sincronizacao abriria a pasta do disco).
' Nunca propaga erro.
Private Sub PptAuto_KeepOpenBooks(d As Object)
    Dim isOpen As Object, w As Workbook, k As Variant
    On Error Resume Next
    Set isOpen = CreateObject("Scripting.Dictionary")
    isOpen.CompareMode = 1
    For Each w In Application.Workbooks
        isOpen(w.FullName) = True
    Next
    For Each k In d.Keys
        If Not isOpen.Exists(CStr(k)) Then d.Remove k
    Next
    Err.Clear
End Sub

' Etiqueta lida SEM engolir o PowerPoint ocupado (PptSync_Tag devolve "" em qualquer erro): ocupado sobe (uma leitura recusada vira nova
' tentativa, nunca "entrada velha"); outro erro ou etiqueta ausente = "". hookPoint: ponto de falha injetada (build Test).
Private Function PptAuto_TagStrict(o As Object, key As String, hookPoint As String) As String
    Dim v As Variant, en As Long, ed As String
    On Error Resume Next
    PptSync_TestPoint hookPoint
    If Err.Number = 0 Then v = o.Tags.Item(key)
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    If Ppt_IsBusyError(en) Then Err.Raise en, "PptAuto_TagStrict", ed
    If en = 0 And Not IsNull(v) And Not IsObject(v) Then PptAuto_TagStrict = CStr(v)
End Function

' A forma e um elemento do ThinQQ (grafico TQQ_CHART, texto vinculado TQQ_TEXT = "1" ou tabela vinculada TQQ_TABLE = "1"; a regra de
' PptSync_ItemKind) sem engolir o ocupado.
Private Function PptAuto_IsItemStrict(shp As Object) As Boolean
    If Len(PptAuto_TagStrict(shp, PPTSYNC_TAG_CHART, "")) > 0 Then PptAuto_IsItemStrict = True: Exit Function
    If PptAuto_TagStrict(shp, PPTTEXT_TAG, "") = "1" Then PptAuto_IsItemStrict = True: Exit Function
    PptAuto_IsItemStrict = (PptAuto_TagStrict(shp, PPTTABLE_TAG, "") = "1")
End Function

' Pasta de origem da forma (como PptSync_ItemBook: vinculo nativo ou TQQ_BOOK), sem engolir o PowerPoint ocupado.
Private Function PptAuto_BookStrict(shp As Object) As String
    Dim lk As String, en As Long, ed As String
    On Error Resume Next
    PptSync_TestPoint "autobook"
    If Err.Number = 0 Then lk = CStr(shp.LinkFormat.SourceFullName)
    en = Err.Number: ed = Err.Description
    On Error GoTo 0
    If Ppt_IsBusyError(en) Then Err.Raise en, "PptAuto_BookStrict", ed
    If en <> 0 Then lk = ""
    PptAuto_BookStrict = PptSync_SourcePath(lk, PptAuto_TagStrict(shp, PPTSYNC_TAG_BOOK, ""))
End Function

' Formas automaticas de pres cuja pasta de origem esta em dirty, pelo indice TQQ_AUTOIDX (nunca varre todas as formas). A pasta da entrada
' pre-filtra ANTES de qualquer chamada ao PowerPoint: entrada de pasta que nao mudou fica no indice como esta, sem achar slide nem forma (a
' entrada velha dela sai na sincronizacao completa, PptSync_RebuildIndex). Entrada da pasta suja: a forma entra se e do ThinQQ
' (PptSync_ItemKind), tem TQQ_AUTO e o vinculo REAL dela aponta para uma pasta de dirty; a mesma forma entra uma vez so. O indice so e
' regravado quando uma entrada sai (velha: slide ou forma apagados, sem TQQ_AUTO, nao e do ThinQQ; ou repetida). PowerPoint ocupado sobe
' (inclusive nas leituras de etiqueta e de vinculo).
Public Function PptAuto_ItemsFor(pres As Object, dirty As Object) As Collection
    Dim out As New Collection, keep As New Collection, e As Variant, sl As Object, shp As Object, removed As Boolean, idx As String, en As Long
    Dim seen As Object, k As String
    Set PptAuto_ItemsFor = out
    idx = PptAuto_TagStrict(pres, PPTSYNC_TAG_INDEX, "autoidx")
    If Len(idx) = 0 Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For Each e In PptIdx_Parse(idx)
        If Not dirty.Exists(CStr(e(2))) Then
            k = "n" & CStr(e(0)) & "|" & LCase$(CStr(e(1)))       ' o PowerPoint acha a forma pelo nome sem diferenciar maiusculas
            If seen.Exists(k) Then removed = True Else seen.Add k, True: keep.Add e
        Else
            Set sl = Nothing: Set shp = Nothing
            On Error Resume Next
            Set sl = pres.Slides.FindBySlideID(CLng(e(0)))
            If Not sl Is Nothing Then Set shp = sl.Shapes(CStr(e(1)))
            en = Err.Number
            Err.Clear
            On Error GoTo 0
            If Ppt_IsBusyError(en) Then Err.Raise en, "PptAuto_ItemsFor", "PowerPoint ocupado"
            If shp Is Nothing Then
                removed = True
            ElseIf PptAuto_TagStrict(shp, PPTSYNC_TAG_AUTO, "autotag") <> "1" Then
                removed = True
            ElseIf Not PptAuto_IsItemStrict(shp) Then          ' = PptSync_ItemKind (grafico, texto ou tabela vinculados), lida sem engolir o ocupado
                removed = True
            Else
                k = PptAuto_ShapeKey(CLng(e(0)), shp, CStr(e(1)))
                If seen.Exists(k) Then
                    removed = True
                Else
                    seen.Add k, True
                    keep.Add e
                    If dirty.Exists(PptAuto_BookStrict(shp)) Then out.Add shp
                End If
            End If
        End If
    Next
    If removed Then
        On Error Resume Next
        If keep.Count = 0 Then PptSync_TagDelete pres, PPTSYNC_TAG_INDEX Else pres.Tags.Add PPTSYNC_TAG_INDEX, PptIdx_Join(keep)
        Err.Clear
        On Error GoTo 0
    End If
End Function

' Chave da forma: SlideID e Shape.Id (objetos de outro processo nao se comparam com Is); sem Id legivel, o nome da entrada sem diferenciar
' maiusculas (como Shapes(nome) do PowerPoint). Nunca propaga erro.
Private Function PptAuto_ShapeKey(slideId As Long, shp As Object, entryName As String) As String
    Dim id As Long
    On Error Resume Next
    id = CLng(shp.Id)
    If Err.Number = 0 Then PptAuto_ShapeKey = CStr(slideId) & "#" & CStr(id) Else PptAuto_ShapeKey = CStr(slideId) & Chr$(31) & LCase$(entryName)
    Err.Clear
End Function

' Nome da apresentacao para o log. Nunca propaga erro.
Private Function PptAuto_PresName(pres As Object) As String
    On Error Resume Next
    PptAuto_PresName = CStr(pres.Name)
    Err.Clear
End Function

' 1a passada do ciclo, numa apresentacao: so ReadOnly, TQQ_AUTOIDX e (havendo entradas da pasta suja) as formas do indice. Somente leitura
' e pulada; com formas a sincronizar acrescenta Array(pres, items) a work. Erro comum: registrado e a apresentacao pulada (as outras
' seguem); PowerPoint ocupado sobe para o ciclo tentar de novo.
Private Sub PptAuto_Collect(pres As Object, work As Collection)
    Dim items As Collection, en As Long, ed As String
    On Error GoTo Fail
    If CLng(pres.ReadOnly) <> 0 Then Exit Sub
    Set items = PptAuto_ItemsFor(pres, mDirty)
    If items.Count > 0 Then work.Add Array(pres, items)
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error GoTo 0                                 ' sem isto o GoTo -1 deixa o Fail armado e o Raise voltaria para ca (laco)
    If Ppt_IsBusyError(en) Then Err.Raise en, "PptAuto_Collect", ed
    PptAuto_Log "PptAuto: apresentacao pulada" & vbTab & PptAuto_PresName(pres) & vbTab & en & vbTab & ed
End Sub

Public Function PptAuto_LastMarkedForTest() As Long
    PptAuto_LastMarkedForTest = mLastMarked
End Function

' Elementos manuais (1.11.0) de uma apresentacao editavel, para a bandeira. Erro comum: registrado e a apresentacao pulada; ocupado sobe.
Private Sub PptAuto_CollectManual(pres As Object, out As Collection)
    Dim items As Collection, shp As Object, en As Long, ed As String
    On Error GoTo Fail
    If CLng(pres.ReadOnly) <> 0 Then Exit Sub
    Set items = PptState_ManualFor(pres, mDirty)
    For Each shp In items
        out.Add shp
    Next
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error GoTo 0
    If Ppt_IsBusyError(en) Then Err.Raise en, "PptAuto_CollectManual", ed
    PptAuto_Log "PptAuto: apresentacao pulada (bandeiras)" & vbTab & PptAuto_PresName(pres) & vbTab & en & vbTab & ed
End Sub

' 2a passada: sincroniza as formas de uma apresentacao (sem forcar; pasta nao aberta nunca e aberta). Registra so quando algo foi atualizado
' ou ha falha/aviso (nao uma linha "sem mudanca" por ciclo). Erro comum: registrado; ocupado sobe.
Private Sub PptAuto_SyncWork(ByVal pres As Object, ByVal items As Collection)
    Dim s As String, en As Long, ed As String
    On Error GoTo Fail
    s = PptSync_Items(items, False, True)
    mLastResult = s
    If Left$(s, Len(PPTAUTO_SUM_NONE)) <> PPTAUTO_SUM_NONE Or InStr(s, vbLf) > 0 Then
        PptAuto_Log "PptAuto" & vbTab & PptAuto_PresName(pres) & vbTab & Replace(s, vbLf, " | ")
    End If
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error GoTo 0                                 ' idem: sem laco no Raise
    If Ppt_IsBusyError(en) Then Err.Raise en, "PptAuto_SyncWork", ed
    PptAuto_Log "PptAuto: apresentacao pulada" & vbTab & PptAuto_PresName(pres) & vbTab & en & vbTab & ed
End Sub

' Guarda o texto da barra de status antes da sincronizacao (PptSync_Items escreve o progresso e devolve a barra ao Excel ao terminar).
Private Sub PptAuto_StatusSave()
    On Error Resume Next
    mSbPrev = Application.StatusBar
    mSbSaved = (Err.Number = 0)
    Err.Clear
End Sub

' Barra de status com o Excel (ninguem escreveu): False, ou o texto "FALSE" que o Excel invisivel devolve (medido em TestDeltaPick).
Private Function PptAuto_StatusIsExcel(v As Variant) As Boolean
    If IsEmpty(v) Or VarType(v) = vbBoolean Then PptAuto_StatusIsExcel = True: Exit Function
    PptAuto_StatusIsExcel = (CStr(v) = "FALSE")
End Function

' Devolve o texto guardado so se a barra esta com o Excel (False) ou com o progresso da sincronizacao: texto escrito por outro recurso
' durante o ciclo fica. Nunca propaga erro.
Private Sub PptAuto_StatusRestore()
    Dim cur As Variant
    If Not mSbSaved Then Exit Sub
    mSbSaved = False
    On Error Resume Next
    cur = Application.StatusBar
    If Not PptAuto_StatusIsExcel(cur) Then
        If Left$(CStr(cur), Len(PPTAUTO_PROGRESS)) <> PPTAUTO_PROGRESS Then mSbPrev = Empty: Exit Sub
    End If
    If PptAuto_StatusIsExcel(mSbPrev) Then Application.StatusBar = False Else Application.StatusBar = CStr(mSbPrev)
    mSbPrev = Empty
    Err.Clear
End Sub

' Alvo do OnTime: sincroniza, nas apresentacoes editaveis do PowerPoint JA aberto, as formas automaticas das pastas que mudaram e continuam
' abertas (sem forcar: a impressao digital pula o que nao mudou). Duas passadas: a 1a so le ReadOnly/TQQ_AUTOIDX (e as formas do indice da
' pasta suja); sem nada a sincronizar a fila esvazia sem mais nenhuma chamada ao PowerPoint (nem a guarda de texto em edicao, nem novas
' tentativas, nem registro). Sem PowerPoint ou opcao desligada: esvazia a fila. Operacao do core com o PowerPoint em andamento (colagem,
' sincronizacao pedida) ou ciclo em andamento: reagenda. Ocupado (ou texto em edicao): nova tentativa em PPTAUTO_DEBOUNCE_S * (n + 1) s,
' ate PPTAUTO_MAX_RETRIES. Um disparo velho (cancelamento que falhou) antes do horario vivo e ignorado sem esquecer o OnTime vivo. Sem caixa
' de mensagem (so log e barra de status, devolvida como estava). Nunca propaga erro. Desde a 1.11.0 tambem marca a bandeira dos manuais
' (PptState_MarkItems: TQQ_STATE pelo indice TQQ_LINKIDX), sem alterar o Excel nem o slide.
Public Sub PptAuto_Tick()
    Dim app As Object, pres As Object, work As Collection, w As Variant, en As Long, ed As String, manual As Collection
    On Error GoTo Fail
    If mDueReal Then
        If Now < mDue - TimeSerial(0, 0, 1) Then Errors_Log "PptAuto: OnTime velho ignorado": Exit Sub
    End If
    mDue = 0: mDueReal = False
    If mDirty Is Nothing Then Exit Sub
    If mDirty.Count = 0 Then Exit Sub
    If mInTick Or Ppt_InOperation() Then PptAuto_Schedule PPTAUTO_DEBOUNCE_S: Exit Sub
    mLastResult = ""
    If Not PptAuto_Enabled() Then PptAuto_ClearQueue: Exit Sub
    PptAuto_KeepOpenBooks mDirty
    If mDirty.Count = 0 Then PptAuto_ClearQueue: Exit Sub
    mInTick = True
    If BUILD_CONFIG = "Test" And Len(mTestNotify) > 0 Then
        Core_Notify mTestNotify, vbExclamation
        Core_Confirm mTestNotify
        mTestNotify = ""
    End If
    If BUILD_CONFIG = "Test" And mTestBusy > 0 Then mTestBusy = mTestBusy - 1: Err.Raise PPTAUTO_ERR_BUSY, "PptAuto_Tick", "ocupado (teste)"
    Set app = Ppt_GetApp(False)
    If Not app Is Nothing Then
        Set work = New Collection
        For Each pres In app.Presentations
            PptAuto_Collect pres, work
        Next
        Set pres = Nothing
        If work.Count > 0 Then
            If PptAuto_PptEditingText(app) Then Err.Raise PPTAUTO_ERR_BUSY, "PptAuto_Tick", "texto em edicao no PowerPoint"
            PptAuto_StatusSave
            For Each w In work
                PptAuto_SyncWork w(0), w(1)
            Next
            PptAuto_StatusRestore
        End If
        ' 1.11.0: os MANUAIS da pasta que mudou ganham (ou perdem) a bandeira TQQ_STATE, sem nada mudar no Excel nem no slide
        Set manual = New Collection
        For Each pres In app.Presentations
            PptAuto_CollectManual pres, manual
        Next
        Set pres = Nothing
        mLastMarked = 0
        If manual.Count > 0 Then mLastMarked = PptState_MarkItems(manual)
    End If
    w = Empty: Set work = Nothing: Set manual = Nothing: Set app = Nothing
    PptAuto_SaveDataBooks mDirty                    ' ainda com mInTick = True: o AfterSave do salvamento e ignorado
    PptAuto_ClearQueue
    mInTick = False
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    w = Empty: Set work = Nothing: Set manual = Nothing: Set pres = Nothing: Set app = Nothing
    PptAuto_StatusRestore
    mInTick = False
    If Ppt_IsBusyError(en) And mRetries < PPTAUTO_MAX_RETRIES Then
        mRetries = mRetries + 1
        PptAuto_Log "PptAuto: PowerPoint ocupado (" & ed & "); nova tentativa " & mRetries
        PptAuto_Schedule PPTAUTO_DEBOUNCE_S * (mRetries + 1)
    Else
        PptAuto_Log "PptAuto: desistiu" & vbTab & en & vbTab & ed
        PptAuto_ClearQueue
    End If
End Sub

' Pastas de DADOS (graficos criados pelo PowerPoint, marca TQQ_DATABOOK) que mudaram: salvas no fim do ciclo, para o arquivo nunca ficar atras
' do slide (spec 14.2). Pastas do usuario nunca sao salvas pelo ThinQQ. O AfterSave cai em PptAuto_OnChange com mInTick = True (ignorado).
Private Sub PptAuto_SaveDataBooks(dirty As Object)
    Dim w As Workbook
    On Error Resume Next
    If dirty Is Nothing Then Exit Sub
    For Each w In Application.Workbooks
        If dirty.Exists(w.FullName) Then
            If Not w.Saved And Not w.ReadOnly And PptData_IsDataBook(w) Then
                Err.Clear
                PptData_SaveQuiet w                                     ' sem caixas do Excel no Save (revisao final M5)
                If Err.Number <> 0 Then PptAuto_Log "PptAuto: pasta de dados nao salva" & vbTab & w.FullName & vbTab & Err.Number & vbTab & Err.Description
            End If
        End If
    Next
    Err.Clear
End Sub
