Attribute VB_Name = "modDeltaPick"
Option Explicit
' Modo "Delta clicando nas barras" (Parte B, spec 2026-10-14 secao 5.1): o botao arma o modo no grafico ativo; o 1o clique numa barra
' escolhe a origem (moldura laranja temporaria tqq_pickmark), o 2o o destino e a seta e inserida (Core_AddDelta). Caminhos do clique:
' Chart.MouseUp + GetChartElement (acha o ponto ja no 1o clique) e Chart.Select (ElementID = xlSeries, Arg1 = serie, Arg2 = ponto; -1
' quando o Excel selecionou a serie inteira). O mesmo clique pelos dois caminhos conta uma vez. Cancela: Esc, clique numa celula, outro
' grafico, outra planilha/pasta ativada, fechar a pasta, o botao de novo, 60 s sem clique (contados do ultimo
' clique aceito: armar ou escolher a origem). Desarmado, os manipuladores saem na primeira
' linha. Estado global mexido SO enquanto armado e devolvido em todo caminho (DeltaPick_Disarm): barra de status, tecla Esc (OnKey),
' OnTime do prazo e do aviso de 4 s (cancelados com o MESMO horario agendado; apontam para o core, nunca para a pasta do usuario) e a
' moldura. Nada aqui propaga erro (todos os pontos de entrada sao chamados por eventos do Excel, OnKey ou OnTime).
Public Const DELTAPICK_MARK As String = "tqq_pickmark"
Public Const DELTAPICK_MSG_FROM As String = "ThinQQ - Delta: clique na barra de ORIGEM (Esc ou um clique numa célula cancela)."
Public Const DELTAPICK_MSG_TO As String = "ThinQQ - Delta: agora clique na barra de DESTINO."
Public Const DELTAPICK_MSG_SAME As String = "ThinQQ - Delta: escolha uma barra diferente da origem."
Public Const DELTAPICK_MSG_NOTBAR As String = "ThinQQ - Delta: clique numa barra (não em rótulos, setas ou linhas auxiliares)."
Public Const DELTAPICK_MSG_SERIES As String = "ThinQQ - Delta: clique de novo na mesma barra para escolher só ela."
Public Const DELTAPICK_MSG_CANCEL As String = "ThinQQ - Delta cancelado."
Public Const DELTAPICK_MSG_OTHERCHART As String = "ThinQQ - Delta cancelado (outro gráfico)."
Public Const DELTAPICK_MSG_TIMEOUT As String = "ThinQQ - Delta cancelado (60 s sem clique)."
Public Const DELTAPICK_MSG_PROTECTED As String = "A planilha está protegida. Desproteja a planilha para inserir o delta."
Public Const DELTAPICK_MSG_ARMFAIL As String = "Não foi possível ativar o delta clicando nas barras neste gráfico. Selecione o gráfico e tente de novo."
Private Const DELTAPICK_TIMEOUT_S As Long = 60
Private Const DELTAPICK_CLEAR_S As Long = 4
Private Const DELTAPICK_DEDUPE_S As Single = 0.6
Private mArmed As Boolean
Private mChartName As String                    ' ChartObject, planilha e pasta do grafico armado
Private mSheetName As String
Private mBookName As String
Private mFromRef As String
Private mLastKey As String                      ' ultimo clique tratado (serie:ponto), de onde veio e quando
Private mLastSrc As String
Private mLastAt As Single
Private mDeadline As Date                       ' fim do prazo de 60 s (o gancho de teste o antecipa sem perder o horario agendado)
Private mTimeoutAt As Date                      ' horario do OnTime do prazo agendado (0 = nenhum): o cancelamento exige o mesmo valor
Private mClearAt As Date                        ' horario do OnTime que apaga o aviso da barra de status (0 = nenhum)
Private mStatusOwned As Boolean                 ' a barra de status mostra um texto nosso (armado ou aviso de 4 s)
Private mTimerFailures As Long                  ' OnTime que nao pode ser agendado/cancelado (testes: deve ficar 0)
Private mSlowMark As Single                     ' gancho de teste: segundos a mais no proximo desenho da moldura (uma vez; 0 = nada)

Public Function DeltaPick_IsArmed() As Boolean
    DeltaPick_IsArmed = mArmed
End Function

Public Function DeltaPick_IsArmedFor(cht As Chart) As Boolean
    On Error Resume Next
    If Not mArmed Then Exit Function
    DeltaPick_IsArmedFor = (cht.Parent.Name = mChartName And cht.Parent.Parent.Name = mSheetName And cht.Parent.Parent.Parent.Name = mBookName)
End Function

Public Function DeltaPick_FromRef() As String
    DeltaPick_FromRef = mFromRef
End Function

' OnTime pendentes do modo (prazo + aviso) e falhas de agendamento/cancelamento. Ganchos de teste (inofensivos em Release).
Public Function DeltaPick_PendingTimers() As Long
    DeltaPick_PendingTimers = Abs(mTimeoutAt <> 0) + Abs(mClearAt <> 0)
End Function

Public Function DeltaPick_TimerFailures() As Long
    DeltaPick_TimerFailures = mTimerFailures
End Function

' Referencia de procedimento do core para OnKey/OnTime (aspas simples do nome escapadas).
Private Function DeltaPick_Proc(proc As String) As String
    DeltaPick_Proc = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & proc
End Function

Private Sub DeltaPick_CancelTimer(ByRef at As Date, proc As String)
    If at = 0 Then Exit Sub
    On Error Resume Next
    Application.OnTime at, DeltaPick_Proc(proc), , False
    If Err.Number <> 0 Then mTimerFailures = mTimerFailures + 1: Errors_Log "DeltaPick_CancelTimer " & proc & vbTab & Err.Number & vbTab & Err.Description
    Err.Clear
    at = 0
End Sub

Private Function DeltaPick_Schedule(secs As Long, proc As String) As Date
    Dim at As Date
    On Error Resume Next
    at = Now + TimeSerial(0, 0, secs)
    Application.OnTime at, DeltaPick_Proc(proc)
    If Err.Number <> 0 Then
        mTimerFailures = mTimerFailures + 1
        Errors_Log "DeltaPick_Schedule " & proc & vbTab & Err.Number & vbTab & Err.Description
        at = 0
    End If
    Err.Clear
    DeltaPick_Schedule = at
End Function

' (Re)comeca o prazo: cancela o OnTime vivo (mesmo horario/procedimento, nada empilha) e agenda outro daqui a secs; o prazo e o horario
' agendado (mDeadline = mTimeoutAt), ou Now + secs se o OnTime nao pode ser agendado.
Private Sub DeltaPick_StartDeadline(secs As Long)
    DeltaPick_CancelTimer mTimeoutAt, "DeltaPick_TimeoutFired"
    mTimeoutAt = DeltaPick_Schedule(secs, "DeltaPick_TimeoutFired")
    If mTimeoutAt <> 0 Then mDeadline = mTimeoutAt Else mDeadline = Now + TimeSerial(0, 0, secs)
End Sub

' Segundos que faltam para o prazo do modo armado (0 desarmado) e para o OnTime agendado (0 sem OnTime). Ganchos de teste (so leitura).
Public Function DeltaPick_SecondsLeft() As Double
    If mArmed Then DeltaPick_SecondsLeft = (mDeadline - Now) * 86400#
End Function

Public Function DeltaPick_TimerSecondsLeft() As Double
    If mTimeoutAt <> 0 Then DeltaPick_TimerSecondsLeft = (mTimeoutAt - Now) * 86400#
End Function

' Gancho de teste: o modo armado ha secs segundos (o OnTime e trocado por um que vence em 60 - secs, como se o tempo tivesse passado).
' Inerte fora do build Test.
Public Sub DeltaPick_AgeForTest(secs As Long)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    If mArmed Then DeltaPick_StartDeadline DELTAPICK_TIMEOUT_S - secs
End Sub

Public Sub DeltaPick_Arm(cht As Chart)
    On Error GoTo Fail
    DeltaPick_Disarm ""
    mChartName = cht.Parent.Name: mSheetName = cht.Parent.Parent.Name: mBookName = cht.Parent.Parent.Parent.Name
    mFromRef = "": mLastKey = "": mLastSrc = ""
    mArmed = True
    mStatusOwned = True
    Application.StatusBar = DELTAPICK_MSG_FROM
    Application.OnKey "{ESC}", DeltaPick_Proc("DeltaPick_EscKey")
    DeltaPick_StartDeadline DELTAPICK_TIMEOUT_S
    Exit Sub
Fail:
    Errors_Log "DeltaPick_Arm" & vbTab & Err.Number & vbTab & Err.Description
    On Error GoTo -1
    DeltaPick_Disarm ""
    Core_Notify DELTAPICK_MSG_ARMFAIL, vbExclamation
End Sub

' Desarma (msg "" = barra de status devolvida ao Excel; senao mostra msg por 4 s). Remove a moldura, devolve o Esc e cancela os OnTime
' pendentes. Sem nada nosso na barra de status e desarmado, nao mexe nela (pode ser de outro programa). Nunca propaga.
Public Sub DeltaPick_Disarm(msg As String)
    Dim co As ChartObject, wasArmed As Boolean
    On Error Resume Next
    wasArmed = mArmed
    mArmed = False
    mFromRef = "": mLastKey = "": mLastSrc = ""
    mSlowMark = 0
    If wasArmed Then
        Set co = Workbooks(mBookName).Worksheets(mSheetName).ChartObjects(mChartName)
        If Not co Is Nothing Then DeltaPick_RemoveMark co.Chart
        Err.Clear
        Application.OnKey "{ESC}"
        Err.Clear
    End If
    DeltaPick_CancelTimer mTimeoutAt, "DeltaPick_TimeoutFired"
    DeltaPick_CancelTimer mClearAt, "DeltaPick_ClearStatus"
    If Len(msg) > 0 Then
        Application.StatusBar = msg
        mStatusOwned = True
        mClearAt = DeltaPick_Schedule(DELTAPICK_CLEAR_S, "DeltaPick_ClearStatus")
    ElseIf wasArmed Or mStatusOwned Then
        Application.StatusBar = False
        mStatusOwned = False
    End If
    Err.Clear
End Sub

' OnTime do aviso: devolve a barra de status ao Excel (se o modo nao foi armado de novo).
Public Sub DeltaPick_ClearStatus()
    On Error Resume Next
    mClearAt = 0                                    ' disparou: nada a cancelar
    If Not mArmed And mStatusOwned Then Application.StatusBar = False: mStatusOwned = False
End Sub

' Esc (OnKey). Desarmado com a tecla ainda ligada (estado do VBA perdido, ex.: projeto reiniciado): devolve o Esc ao Excel.
Public Sub DeltaPick_EscKey()
    On Error Resume Next
    If mArmed Then
        DeltaPick_Disarm DELTAPICK_MSG_CANCEL
    Else
        Application.OnKey "{ESC}"
    End If
End Sub

' O OnTime do prazo disparou. Armado: desarma se o prazo do modo atual chegou, com 2 s de folga (o Date exato pode ficar um ulp antes e o
' modo ficaria armado com o Esc preso); mais cedo que isso o disparo e de um OnTime velho (um cancelamento que falhou, o VBA reiniciado) e
' e ignorado sem esquecer o OnTime vivo. Desarmado: so devolve o Esc.
Public Sub DeltaPick_TimeoutFired()
    On Error Resume Next
    If mArmed Then
        If Now < mDeadline - TimeSerial(0, 0, 2) Then
            Errors_Log "DeltaPick: OnTime do prazo velho ignorado (faltam " & Format$((mDeadline - Now) * 86400#, "0") & " s)"
            Exit Sub
        End If
        mTimeoutAt = 0                              ' disparou: nada a cancelar
        DeltaPick_Disarm DELTAPICK_MSG_TIMEOUT
    Else
        mTimeoutAt = 0
        Application.OnKey "{ESC}"
    End If
End Sub

' Verificacao do prazo pelo relogio (gancho de teste DeltaPick_ExpireForTest; o OnTime real chama DeltaPick_TimeoutFired).
Public Sub DeltaPick_Timeout()
    On Error Resume Next
    If mArmed And Now >= mDeadline Then DeltaPick_Disarm DELTAPICK_MSG_TIMEOUT
End Sub

' Gancho de teste: simula o disparo REAL do OnTime do prazo (cancela o OnTime real, como se o Excel o tivesse consumido, poe o prazo 1 s
' adiante, como um Date que chega um pouco antes do horario, e chama o manipulador do disparo). Inerte fora do build Test.
Public Sub DeltaPick_FireTimeoutForTest()
    If BUILD_CONFIG <> "Test" Then Exit Sub
    DeltaPick_CancelTimer mTimeoutAt, "DeltaPick_TimeoutFired"
    If mArmed Then mDeadline = Now + TimeSerial(0, 0, 1)
    DeltaPick_TimeoutFired
End Sub

' Gancho de teste: o proximo desenho da moldura demora secs a mais (grafico pesado). Vale uma vez; inerte fora do build Test.
Public Sub DeltaPick_SlowMarkForTest(secs As Single)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mSlowMark = secs
End Sub

' Gancho de teste: faz o prazo do modo armado vencer agora (o OnTime real continua agendado e e cancelado no desarme). Inerte fora do Test.
Public Sub DeltaPick_ExpireForTest()
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mDeadline = Now - TimeSerial(0, 0, 1)
End Sub

' Clique numa celula (App_SheetSelectionChange).
Public Sub DeltaPick_OnSheetClick()
    On Error Resume Next
    If mArmed Then DeltaPick_Disarm DELTAPICK_MSG_CANCEL
End Sub

' Outra planilha ou pasta ativada (App_SheetActivate / App_WorkbookActivate): os observadores mudam de planilha e o grafico armado deixa de
' receber cliques. bookName/sheetName = a que ficou ativa ("" em sheetName = so a pasta).
Public Sub DeltaPick_OnActivate(bookName As String, sheetName As String)
    On Error Resume Next
    If Not mArmed Then Exit Sub
    If bookName <> mBookName Or (Len(sheetName) > 0 And sheetName <> mSheetName) Then DeltaPick_Disarm DELTAPICK_MSG_CANCEL
End Sub

' A pasta do grafico armado vai fechar (App_WorkbookBeforeClose): desarma antes (a moldura sai antes da pergunta de salvar).
Public Sub DeltaPick_OnBookClose(bookName As String)
    On Error Resume Next
    If mArmed And bookName = mBookName Then DeltaPick_Disarm ""
End Sub

' Chart.Select de QUALQUER grafico observado. Outro grafico cancela; rotulo, linhas da serie, barras de erro, formas (setas, a moldura):
' "clique numa barra" e segue armado; area, eixo, legenda, titulo: segue armado em silencio.
Public Sub DeltaPick_OnSelect(cht As Chart, elementId As Long, a1 As Long, a2 As Long)
    If Not mArmed Then Exit Sub
    On Error Resume Next
    If Not DeltaPick_IsArmedFor(cht) Then DeltaPick_Disarm DELTAPICK_MSG_OTHERCHART: Exit Sub
    If elementId <> xlSeries Then
        Select Case elementId
            Case xlDataLabel, xlTrendline, xlErrorBars, xlXErrorBars, xlYErrorBars, xlShape, xlSeriesLines, xlHiLoLines, xlDropLines, xlLeaderLines
                Application.StatusBar = DELTAPICK_MSG_NOTBAR
        End Select
        Exit Sub
    End If
    If a2 < 1 Then
        If Not DeltaPick_IsDuplicate(a1 & ":", "select") Then Application.StatusBar = DELTAPICK_MSG_SERIES
        Exit Sub
    End If
    DeltaPick_Take cht, a1, a2, "select"
End Sub

' Chart.MouseUp: acha o elemento sob o ponteiro (coordenadas do proprio evento, as mesmas do GetChartElement).
Public Sub DeltaPick_OnMouseUp(cht As Chart, x As Long, y As Long)
    Dim id As Long, a1 As Long, a2 As Long
    If Not mArmed Then Exit Sub
    On Error Resume Next
    If Not DeltaPick_IsArmedFor(cht) Then Exit Sub
    cht.GetChartElement x, y, id, a1, a2
    If Err.Number <> 0 Then Err.Clear: Exit Sub
    If id = xlSeries And a2 >= 1 Then DeltaPick_Take cht, a1, a2, "mouse"
End Sub

' O mesmo clique chega pelos dois caminhos (MouseUp e Select): mesma chave, OUTRA fonte, em menos de 0,6 s = repetido.
Private Function DeltaPick_IsDuplicate(key As String, src As String) As Boolean
    Dim t As Single
    t = Timer
    If Len(mLastKey) > 0 And Left$(mLastKey, Len(key)) = key And src <> mLastSrc And Abs(t - mLastAt) < DELTAPICK_DEDUPE_S Then DeltaPick_IsDuplicate = True
End Function

' Clique numa barra (serie s, ponto p) pelo caminho src. Erro inesperado: registra e desarma (nada fica pendurado).
Private Sub DeltaPick_Take(cht As Chart, s As Long, p As Long, src As String)
    Dim ref As String, key As String, fromRef As String
    On Error GoTo Fail
    key = s & ":" & p
    If key = mLastKey And src <> mLastSrc And Abs(Timer - mLastAt) < DELTAPICK_DEDUPE_S Then Exit Sub
    ' MouseUp logo depois de um Select ACEITO: e o mesmo clique (o hit test do proprio Excel ja o contou), qualquer que seja o ponto; senao
    ' um clique na borda (zoom/DPI) com os dois caminhos discordando completaria o delta. Select -> Select e MouseUp -> MouseUp seguem valendo.
    If src = "mouse" And mLastSrc = "select" And Abs(Timer - mLastAt) < DELTAPICK_DEDUPE_S Then
        If key <> mLastKey Then Errors_Log "DeltaPick: MouseUp " & key & " ignorado logo depois do Select " & mLastKey & " (hit test diferente: zoom/DPI?)"
        Exit Sub
    End If
    mLastKey = key: mLastSrc = src: mLastAt = Timer
    If Not Delta_ResolvePick(cht, s, p, ref) Then Application.StatusBar = DELTAPICK_MSG_NOTBAR: Exit Sub
    If Len(mFromRef) = 0 Then
        mFromRef = ref
        DeltaPick_DrawMark cht, ref
        Application.StatusBar = DELTAPICK_MSG_TO
        DeltaPick_StartDeadline DELTAPICK_TIMEOUT_S      ' os 60 s contam do ultimo clique aceito
        mLastAt = Timer                                 ' a janela do mesmo clique (0,6 s) comeca depois do desenho da moldura
        Exit Sub
    End If
    If ref = mFromRef Then Application.StatusBar = DELTAPICK_MSG_SAME: Exit Sub
    fromRef = mFromRef
    DeltaPick_Disarm ""
    Core_AddDelta fromRef, ref, "both", cht
    Exit Sub
Fail:
    Errors_Log "DeltaPick_Take" & vbTab & Err.Number & vbTab & Err.Description
    On Error GoTo -1
    DeltaPick_Disarm ""
End Sub

' Moldura laranja de 2 pt em volta da barra de origem (cosmetica: falha so registra). Sentido dos valores pelo eixo da serie (Delta_ValCoord:
' a linha da combinacao no eixo secundario), preso a area interna.
Private Sub DeltaPick_DrawMark(cht As Chart, ref As String)
    Dim v As Double, lo As Double, hi As Double, c As Double, half As Double, a As Double, b As Double, m As Shape, lft As Double, t As Single
    On Error GoTo Done
    If BUILD_CONFIG = "Test" And mSlowMark > 0 Then
        t = Timer
        Do While Abs(Timer - t) < mSlowMark And Abs(Timer - t) < 5
        Loop
    End If
    mSlowMark = 0
    DeltaPick_RemoveMark cht
    If Not Delta_Extent(cht, ref, v, lo, hi) Then Exit Sub
    c = Delta_CatCenter(cht, ref)
    half = Delta_BarHalf(cht, ref) + 2
    a = Delta_ValCoord(cht, ref, hi): b = Delta_ValCoord(cht, ref, lo)
    If a < b Then lft = a Else lft = b
    If Geom_ChartIsHorizontal(cht) Then
        Set m = cht.Shapes.AddShape(msoShapeRectangle, lft - 2, c - half, Abs(b - a) + 4, 2 * half)
    Else
        Set m = cht.Shapes.AddShape(msoShapeRectangle, c - half, lft - 2, 2 * half, Abs(b - a) + 4)
    End If
    m.Name = DELTAPICK_MARK
    m.AlternativeText = "kind=pickmark"
    m.Fill.Visible = msoFalse
    m.Line.ForeColor.RGB = RGB(237, 125, 49)
    m.Line.Weight = 2
    Exit Sub
Done:
    Errors_Log "DeltaPick_DrawMark" & vbTab & Err.Number & vbTab & Err.Description
End Sub

Public Sub DeltaPick_RemoveMark(cht As Chart)
    On Error Resume Next
    cht.Shapes(DELTAPICK_MARK).Delete
    Err.Clear
End Sub
