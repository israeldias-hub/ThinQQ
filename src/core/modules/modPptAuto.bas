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
Private Const PPTAUTO_ERR_BUSY As Long = -2147418111        ' 0x80010001 (RPC_E_CALL_REJECTED): adiar como PowerPoint ocupado
Private mTestEnabled As Variant                 ' gancho: Empty = padrao
Private mTestNewAuto As Variant
Private mDue As Date                                        ' horario do OnTime agendado (0 = nenhum): o cancelamento exige o mesmo valor
Private mDueReal As Boolean                                 ' mDue tem um Application.OnTime de verdade (senao so o horario, gancho de teste)
Private mDirty As Object                                    ' Scripting.Dictionary (texto): FullName das pastas que mudaram
Private mInTick As Boolean                                  ' reentrada e eventos do proprio ciclo
Private mRetries As Long
Private mLastResult As String
Private mTimerFailures As Long                              ' OnTime que nao pode ser agendado/cancelado (testes: deve ficar 0)
Private mTestRealTimer As Boolean                           ' gancho: build Test so agenda OnTime real quando o teste pede
Private mTestBusy As Long                                   ' gancho: quantos ciclos simulam PowerPoint ocupado

Public Sub PptAuto_SetFlagsForTest(enabled As Variant, newAuto As Variant)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestEnabled = enabled
    mTestNewAuto = newAuto
End Sub

Public Sub PptAuto_ResetTestHooks()
    mTestEnabled = Empty
    mTestNewAuto = Empty
    mTestBusy = 0
    mTimerFailures = 0
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

Public Function PptAuto_PendingCountForTest() As Long
    If mDirty Is Nothing Then Exit Function
    PptAuto_PendingCountForTest = mDirty.Count
End Function

Public Function PptAuto_IsScheduledForTest() As Boolean
    PptAuto_IsScheduledForTest = (mDue <> 0)
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

' Pasta que pode ter elementos no PowerPoint: salva (o vinculo aponta para o arquivo) e com a planilha _ThinQQ (graficos do ThinQQ). As
' versoes 1.8.0/1.9.0 acrescentam os nomes definidos TQQ_X_/TQQ_T_. Nunca propaga erro.
Public Function PptAuto_IsCandidate(wb As Workbook) As Boolean
    Dim ws As Object
    On Error Resume Next
    If wb Is Nothing Then Exit Function
    If Len(wb.Path) = 0 Then Exit Function
    Set ws = wb.Worksheets(CALC_SHEET)
    PptAuto_IsCandidate = Not ws Is Nothing
    Err.Clear
End Function

' Eventos do Excel (SheetChange, SheetCalculate, WorkbookAfterSave): marca a pasta e (re)agenda o ciclo para daqui a PPTAUTO_DEBOUNCE_S.
' Barato quando nada se aplica (pasta sem arquivo ou sem ThinQQ sai antes de ler a opcao). Ignora os eventos do proprio ciclo. Nunca
' propaga erro.
Public Sub PptAuto_OnChange(wb As Workbook)
    On Error GoTo Done
    If mInTick Then Exit Sub
    If Not PptAuto_IsCandidate(wb) Then Exit Sub
    If Not PptAuto_Enabled() Then Exit Sub
    If mDirty Is Nothing Then
        Set mDirty = CreateObject("Scripting.Dictionary")
        mDirty.CompareMode = 1
    End If
    mDirty(wb.FullName) = True
    mRetries = 0
    PptAuto_Schedule PPTAUTO_DEBOUNCE_S
Done:
End Sub

Private Function PptAuto_Proc() As String
    PptAuto_Proc = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PptAuto_Tick"
End Function

' Agenda o ciclo para daqui a secs, cancelando antes o OnTime vivo (um so de cada vez). Sem OnTime possivel: registra e fica sem agenda (a
' proxima mudanca tenta de novo). Nunca propaga erro.
Private Sub PptAuto_Schedule(secs As Long)
    Dim real_ As Boolean
    PptAuto_CancelTimer
    mDue = Now + TimeSerial(0, 0, secs)
    real_ = True
    If BUILD_CONFIG = "Test" Then real_ = mTestRealTimer
    If Not real_ Then Exit Sub
    On Error Resume Next
    Application.OnTime mDue, PptAuto_Proc()
    If Err.Number <> 0 Then
        mTimerFailures = mTimerFailures + 1
        Errors_Log "PptAuto_Schedule" & vbTab & Err.Number & vbTab & Err.Description
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
        If Err.Number <> 0 Then mTimerFailures = mTimerFailures + 1: Errors_Log "PptAuto_CancelTimer" & vbTab & Err.Number & vbTab & Err.Description
        Err.Clear
    End If
    mDue = 0
    mDueReal = False
End Sub

' Desliga tudo (Core_OnUnload, testes): desagenda e esvazia a fila. Nunca propaga erro.
Public Sub PptAuto_Cancel()
    On Error Resume Next
    PptAuto_CancelTimer
    Set mDirty = Nothing
    mRetries = 0
    mInTick = False
    mTestBusy = 0
    mTestRealTimer = False
    mLastResult = ""
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

' Formas automaticas de pres cuja pasta de origem esta em dirty, pelo indice TQQ_AUTOIDX (nunca varre todas as formas). A pasta da entrada
' so pre-filtra: a forma entra se o vinculo REAL dela (PptSync_ItemBook) aponta para uma pasta de dirty; a mesma forma entra uma vez so.
' Entradas velhas (slide ou forma apagados, TQQ_AUTO ausente) saem do indice. PowerPoint ocupado sobe.
Public Function PptAuto_ItemsFor(pres As Object, dirty As Object) As Collection
    Dim out As New Collection, keep As New Collection, e As Variant, sl As Object, shp As Object, stale As Boolean, idx As String, en As Long
    Dim seen As Object, k As String
    Set PptAuto_ItemsFor = out
    idx = PptSync_Tag(pres, PPTSYNC_TAG_INDEX)
    If Len(idx) = 0 Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For Each e In PptIdx_Parse(idx)
        Set sl = Nothing: Set shp = Nothing
        On Error Resume Next
        Set sl = pres.Slides.FindBySlideID(CLng(e(0)))
        If Not sl Is Nothing Then Set shp = sl.Shapes(CStr(e(1)))
        en = Err.Number
        Err.Clear
        On Error GoTo 0
        If Ppt_IsBusyError(en) Then Err.Raise en, "PptAuto_ItemsFor", "PowerPoint ocupado"
        If shp Is Nothing Then
            stale = True
        ElseIf PptSync_Tag(shp, PPTSYNC_TAG_AUTO) <> "1" Then
            stale = True
        Else
            keep.Add e
            If dirty.Exists(CStr(e(2))) Then
                If dirty.Exists(PptSync_ItemBook(shp)) Then
                    k = PptAuto_ShapeKey(CLng(e(0)), shp, CStr(e(1)))
                    If Not seen.Exists(k) Then seen.Add k, True: out.Add shp
                End If
            End If
        End If
    Next
    If stale Then
        On Error Resume Next
        If keep.Count = 0 Then PptSync_TagDelete pres, PPTSYNC_TAG_INDEX Else pres.Tags.Add PPTSYNC_TAG_INDEX, PptIdx_Join(keep)
        Err.Clear
    End If
End Function

' Chave da forma: SlideID e Shape.Id (objetos de outro processo nao se comparam com Is); sem Id legivel, o nome da entrada. Nunca propaga erro.
Private Function PptAuto_ShapeKey(slideId As Long, shp As Object, entryName As String) As String
    Dim id As Long
    On Error Resume Next
    id = CLng(shp.Id)
    If Err.Number = 0 Then PptAuto_ShapeKey = CStr(slideId) & "#" & CStr(id) Else PptAuto_ShapeKey = CStr(slideId) & Chr$(31) & entryName
    Err.Clear
End Function

' Nome da apresentacao para o log. Nunca propaga erro.
Private Function PptAuto_PresName(pres As Object) As String
    On Error Resume Next
    PptAuto_PresName = CStr(pres.Name)
    Err.Clear
End Function

' Uma apresentacao do ciclo: somente leitura e pulada; erro comum e registrado e a apresentacao pulada (as outras seguem); PowerPoint
' ocupado sobe para o ciclo tentar de novo.
Private Sub PptAuto_SyncPres(pres As Object)
    Dim items As Collection, s As String, en As Long, ed As String
    On Error GoTo Fail
    If CLng(pres.ReadOnly) <> 0 Then Exit Sub
    Set items = PptAuto_ItemsFor(pres, mDirty)
    If items.Count = 0 Then Exit Sub
    s = PptSync_Items(items, False)
    mLastResult = s
    Errors_Log "PptAuto" & vbTab & PptAuto_PresName(pres) & vbTab & Replace(s, vbLf, " | ")
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error GoTo 0                                 ' sem isto o GoTo -1 deixa o Fail armado e o Raise voltaria para ca (laco)
    If Ppt_IsBusyError(en) Then Err.Raise en, "PptAuto_SyncPres", ed
    Errors_Log "PptAuto: apresentacao pulada" & vbTab & PptAuto_PresName(pres) & vbTab & en & vbTab & ed
End Sub

' Alvo do OnTime: sincroniza, nas apresentacoes editaveis do PowerPoint JA aberto, as formas automaticas das pastas que mudaram e continuam
' abertas (sem forcar: a impressao digital pula o que nao mudou). Sem PowerPoint ou opcao desligada: esvazia a fila. Ocupado (ou texto em
' edicao no PowerPoint): nova tentativa em PPTAUTO_DEBOUNCE_S * (n + 1) s, ate PPTAUTO_MAX_RETRIES. Um disparo velho (um cancelamento que
' falhou) antes do horario vivo e ignorado sem esquecer o OnTime vivo. Sem caixa de mensagem (so log e barra de status). Nunca propaga erro.
Public Sub PptAuto_Tick()
    Dim app As Object, pres As Object, en As Long, ed As String
    On Error GoTo Fail
    If mDueReal Then
        If Now < mDue - TimeSerial(0, 0, 1) Then Errors_Log "PptAuto: OnTime velho ignorado": Exit Sub
    End If
    mDue = 0: mDueReal = False
    If mDirty Is Nothing Then Exit Sub
    If mDirty.Count = 0 Then Exit Sub
    If mInTick Then PptAuto_Schedule PPTAUTO_DEBOUNCE_S: Exit Sub
    mLastResult = ""
    If Not PptAuto_Enabled() Then mDirty.RemoveAll: mRetries = 0: Exit Sub
    PptAuto_KeepOpenBooks mDirty
    If mDirty.Count = 0 Then mRetries = 0: Exit Sub
    mInTick = True
    If BUILD_CONFIG = "Test" And mTestBusy > 0 Then mTestBusy = mTestBusy - 1: Err.Raise PPTAUTO_ERR_BUSY, "PptAuto_Tick", "ocupado (teste)"
    Set app = Ppt_GetApp(False)
    If Not app Is Nothing Then
        If PptAuto_PptEditingText(app) Then Err.Raise PPTAUTO_ERR_BUSY, "PptAuto_Tick", "texto em edicao no PowerPoint"
        For Each pres In app.Presentations
            PptAuto_SyncPres pres
        Next
    End If
    Set pres = Nothing: Set app = Nothing
    mDirty.RemoveAll
    mRetries = 0
    mInTick = False
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Set pres = Nothing: Set app = Nothing
    mInTick = False
    If Ppt_IsBusyError(en) And mRetries < PPTAUTO_MAX_RETRIES Then
        mRetries = mRetries + 1
        Errors_Log "PptAuto: PowerPoint ocupado (" & ed & "); nova tentativa " & mRetries
        PptAuto_Schedule PPTAUTO_DEBOUNCE_S * (mRetries + 1)
    Else
        Errors_Log "PptAuto: desistiu" & vbTab & en & vbTab & ed
        If Not mDirty Is Nothing Then mDirty.RemoveAll
        mRetries = 0
    End If
End Sub
