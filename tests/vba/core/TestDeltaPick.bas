Attribute VB_Name = "TestDeltaPick"
Option Explicit
' Parte B (1.5.0): modo "Delta clicando nas barras". Eventos REAIS por Points(i).Select (medido: dispara Chart_Select(3, serie, ponto));
' o caminho do MouseUp e chamado com pixels calculados (pt x 4/3 a 100%) depois de conferir GetChartElement. Todo teste desarma o modo,
' devolve a barra de status e desliga os observadores tambem quando falha (FailTeardown): nada do modo vaza para o teste seguinte.

Private Sub Teardown(wb As Workbook)
    DeltaPick_Disarm ""
    Application.StatusBar = False
    Watchers_Stop
    CloseTemp wb
End Sub

' Caminho de falha de um teste: guarda o erro, desfaz tudo (modo, status, observadores, pasta) e levanta de novo para o corredor.
Private Sub FailTeardown(wb As Workbook)
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    Teardown wb
    On Error GoTo 0
    Err.Raise en, "TestDeltaPick", ed
End Sub

Private Function Fixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Tri", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array("T1", 10, 5)
    ws.Range("A3:C3").Value = Array("T2", 20, 6)
    ws.Range("A4:C4").Value = Array("T3", 30, -4)
    Set Fixture = ws.Range("A1:C4")
End Function

Private Function ArmedChart(wb As Workbook, kind As String) As Chart
    Dim rng As Range, c As Chart
    Set rng = Fixture(wb)
    rng.Select
    Core_Insert kind
    Set c = Core_ActiveChart()
    c.Parent.Activate
    Core_DeltaPick
    Set ArmedChart = c
End Function

Private Function HasShape(c As Chart, nm As String) As Boolean
    Dim s As Shape
    On Error Resume Next
    Set s = c.Shapes(nm)
    HasShape = Not s Is Nothing
End Function

Private Function StatusText() As String
    ' Barra de status devolvida ao Excel: o Excel invisivel do teste le False como o texto "FALSE" (medido), nao como Boolean.
    Dim v As Variant
    v = Application.StatusBar
    If VarType(v) = vbBoolean Then StatusText = "": Exit Function
    StatusText = CStr(v)
    If StatusText = "FALSE" Then StatusText = ""
End Function

Public Sub Test_DeltaPick_TwoClicksBySelectEvents()
    Dim wb As Workbook, c As Chart
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    AssertTrue DeltaPick_IsArmedFor(c), "armado no grafico ativo"
    AssertEqual DELTAPICK_MSG_FROM, StatusText()
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(1).Select
    AssertEqual "1:1", DeltaPick_FromRef()
    AssertEqual DELTAPICK_MSG_TO, StatusText()
    AssertTrue HasShape(c, DELTAPICK_MARK), "moldura na origem"
    c.SeriesCollection(Delta_SeriesIndex(c, 2)).Points(3).Select
    AssertFalse DeltaPick_IsArmed(), "desarmado depois do destino"
    AssertTrue HasShape(c, "tqq_ann_delta_1")
    AssertEqual "kind=delta;v=1;from=1:1;to=2:3;fmt=both", c.Shapes("tqq_ann_delta_1").AlternativeText
    AssertFalse HasShape(c, DELTAPICK_MARK), "moldura removida"
    AssertEqual "", StatusText(), "barra de status devolvida"
    AssertEqual 0, DeltaPick_PendingTimers(), "nenhum OnTime pendente depois da insercao"
    AssertEqual 0, DeltaPick_TimerFailures(), "todo OnTime cancelado com o mesmo horario"
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

Public Sub Test_DeltaPick_SeriesSelectionAsksAgain_SameBar_NotBar()
    Dim wb As Workbook, c As Chart, s As Long, i As Long
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Select
    AssertEqual DELTAPICK_MSG_SERIES, StatusText(), "serie inteira (1o clique do Excel): pede o 2o clique"
    AssertEqual "", DeltaPick_FromRef()
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(2).Select
    AssertEqual "1:2", DeltaPick_FromRef()
    ' selecionar de novo o ponto JA selecionado nao dispara Select no Excel (medido); o clique repetido chega como o Excel o mandaria:
    DeltaPick_OnSelect c, xlSeries, Delta_SeriesIndex(c, 1), 2
    AssertEqual DELTAPICK_MSG_SAME, StatusText(), "mesma barra (pelo mesmo caminho) nao vira destino"
    AssertTrue DeltaPick_IsArmed()
    c.ChartArea.Select
    AssertTrue DeltaPick_IsArmed(), "clique na area do grafico: segue armado"
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(2).Select
    AssertEqual DELTAPICK_MSG_SAME, StatusText(), "mesma barra por evento real (depois de outro elemento)"
    AssertEqual "1:2", DeltaPick_FromRef()
    Teardown wb
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "stacked")
    For i = 1 To c.SeriesCollection.Count
        If Style_IsHelperSeries(c.SeriesCollection(i)) Then s = i: Exit For
    Next
    If s > 0 Then
        c.SeriesCollection(s).Points(1).Select
        AssertEqual DELTAPICK_MSG_NOTBAR, StatusText(), "auxiliar tqq_*: nao e barra"
        AssertEqual "", DeltaPick_FromRef()
    Else
        Errors_Log "TestDeltaPick: sem serie auxiliar na fixture empilhada (trecho pulado)"
    End If
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' 1o clique: MouseUp + GetChartElement acha o ponto; o mesmo clique chegando pelo Select nao conta de novo.
Public Sub Test_DeltaPick_MouseUpFindsPoint_AndDedupe()
    Dim wb As Workbook, c As Chart, pt As Point, x As Long, y As Long, id As Long, a1 As Long, a2 As Long, s As Long
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    ForceLayout c
    s = Delta_SeriesIndex(c, 2)
    Set pt = c.SeriesCollection(s).Points(2)
    x = CLng((pt.Left + pt.Width / 2) * 4 / 3): y = CLng((pt.Top + pt.Height / 2) * 4 / 3)
    c.GetChartElement x, y, id, a1, a2
    If id <> xlSeries Or a1 <> s Or a2 <> 2 Then
        Errors_Log "TestDeltaPick: GetChartElement(" & x & "," & y & ") deu " & id & "/" & a1 & "/" & a2 & " (pixels diferentes de pt x 4/3 aqui): trecho do MouseUp pulado"
    Else
        DeltaPick_OnMouseUp c, x, y
        AssertEqual "2:2", DeltaPick_FromRef(), "o MouseUp ja escolhe a barra"
        c.SeriesCollection(s).Points(2).Select
        AssertEqual DELTAPICK_MSG_TO, StatusText(), "o Select do mesmo clique nao vira 'mesma barra'"
        ' clique fora das barras (area do grafico, canto 2,2 px): o MouseUp nao muda nada
        DeltaPick_OnMouseUp c, 2, 2
        AssertEqual "2:2", DeltaPick_FromRef(), "clique no vazio: segue com a origem"
        AssertTrue DeltaPick_IsArmed(), "clique no vazio: segue armado"
    End If
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

Public Sub Test_DeltaPick_CancelPaths()
    Dim wb As Workbook, c As Chart, c2 As Chart, rng As Range
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    Core_DeltaPick
    AssertFalse DeltaPick_IsArmed(), "o botao de novo cancela"
    AssertEqual DELTAPICK_MSG_CANCEL, StatusText()
    Core_DeltaPick
    DeltaPick_EscKey
    AssertFalse DeltaPick_IsArmed(), "Esc cancela"
    Core_DeltaPick
    DeltaPick_ExpireForTest
    DeltaPick_Timeout
    AssertFalse DeltaPick_IsArmed(), "60 s sem clique cancela"
    AssertEqual 0, DeltaPick_TimerFailures(), "o tempo do teste nao perde o OnTime real (cancelado com o horario agendado)"
    Watchers_Start
    c.Parent.Activate
    Core_DeltaPick
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(1).Select
    AssertTrue HasShape(c, DELTAPICK_MARK)
    wb.Worksheets(1).Range("H20").Select
    AssertFalse DeltaPick_IsArmed(), "clique numa celula cancela"
    AssertFalse HasShape(c, DELTAPICK_MARK), "moldura removida no cancelamento"
    Set rng = wb.Worksheets(1).Range("A1:C4")
    rng.Select
    Core_Insert "clustered"
    Set c2 = Core_ActiveChart()
    c.Parent.Activate
    Core_DeltaPick
    c2.SeriesCollection(1).Points(1).Select
    AssertFalse DeltaPick_IsArmed(), "clique noutro grafico cancela"
    DeltaPick_Disarm ""
    AssertEqual 0, DeltaPick_PendingTimers(), "nenhum OnTime pendente (o aviso de 4 s tambem e cancelado)"
    AssertEqual 0, DeltaPick_TimerFailures()
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

Public Sub Test_DeltaPick_NotArmed_NoEffect_AndRefusals()
    Dim wb As Workbook, rng As Range, c As Chart, i As Long, t As Single, ws As Worksheet
    On Error GoTo Fail
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "clustered"
    Set c = Core_ActiveChart()
    c.Parent.Activate
    t = Timer
    For i = 1 To 200
        c.SeriesCollection(1).Points(1 + (i Mod 3)).Select
    Next
    t = Timer - t
    Errors_Log "TestDeltaPick: 200 selecoes desarmadas em " & Format$(t, "0.00") & " s"
    ' limite folgado (3 s): pega um manipulador que faz trabalho desarmado sem falhar por uma maquina lenta ou ocupada (B6, revisao da B3)
    AssertTrue t < 3, "desarmado: eventos baratos (200 selecoes em " & t & " s; medido na spec: 0,62 s)"
    DeltaPick_OnMouseUp c, 10, 10
    AssertFalse HasShape(c, "tqq_ann_delta_1"), "desarmado: nada inserido"
    AssertFalse HasShape(c, DELTAPICK_MARK), "desarmado: sem moldura"
    AssertEqual "", StatusText()
    Teardown wb
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1").Value = "Cat": ws.Range("B1").Value = "V": ws.Range("A2").Value = "a": ws.Range("B2").Value = 1: ws.Range("A3").Value = "b": ws.Range("B3").Value = 2
    ws.Range("A1:B3").Select
    Core_Insert "pie"
    Core_ClearLastNotice
    Core_DeltaPick
    AssertEqual ANN_DELTA_REFUSAL, Core_LastNotice()
    AssertFalse DeltaPick_IsArmed()
    Teardown wb
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "clustered"
    Set c = Core_ActiveChart()
    c.Parent.Locked = False                 ' grafico bloqueado numa planilha protegida nem pode ser ativado (medido: "formas bloqueadas para selecao")
    rng.Worksheet.Protect DrawingObjects:=True
    c.Parent.Activate
    AssertTrue Not Core_ActiveChart() Is Nothing, "grafico ativo na planilha protegida"
    Core_ClearLastNotice
    Core_DeltaPick
    AssertEqual DELTAPICK_MSG_PROTECTED, Core_LastNotice()
    AssertFalse DeltaPick_IsArmed()
    rng.Worksheet.Unprotect
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' Grafico colado/copiado ainda sem observador (Watchers so religa na troca de planilha): o botao religa antes de armar.
Public Sub Test_DeltaPick_UnwatchedChart_IsHookedOnArm()
    Dim wb As Workbook, c As Chart
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Fixture(wb).Select
    Core_Insert "clustered"
    Set c = Core_ActiveChart()
    Watchers_Stop
    AssertEqual 0, Watchers_Count()
    c.Parent.Activate
    Core_DeltaPick
    AssertTrue DeltaPick_IsArmedFor(c)
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(1).Select
    AssertEqual "1:1", DeltaPick_FromRef(), "o evento chega depois de armar"
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' Moldura: envolve a barra de origem (geometria do Excel pelas Points, independente da conta de producao), tambem nas barras horizontais.
Public Sub Test_DeltaPick_MarkWrapsOriginBar()
    Dim wb As Workbook, c As Chart, kinds As Variant, j As Long, pt As Point, m As Shape, k As String
    On Error GoTo Fail
    kinds = Array("clustered", "barclustered")
    For j = 0 To 1
        Set wb = NewTempWorkbook()
        k = CStr(kinds(j))
        Set c = ArmedChart(wb, k)
        ForceLayout c
        c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(2).Select
        AssertEqual "1:2", DeltaPick_FromRef(), k
        AssertTrue HasShape(c, DELTAPICK_MARK), k & ": moldura"
        Set m = c.Shapes(DELTAPICK_MARK)
        Set pt = c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(2)
        AssertTrue m.Left <= pt.Left + 1, k & ": esquerda " & m.Left & " vs " & pt.Left
        AssertTrue m.Top <= pt.Top + 1, k & ": topo " & m.Top & " vs " & pt.Top
        AssertTrue m.Left + m.Width >= pt.Left + pt.Width - 1, k & ": direita"
        AssertTrue m.Top + m.Height >= pt.Top + pt.Height - 1, k & ": base"
        AssertTrue m.Width <= pt.Width + 12 And m.Height <= pt.Height + 12, k & ": moldura justa"
        AssertEqual "kind=pickmark", m.AlternativeText
        Teardown wb
        Set wb = Nothing
    Next
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' Cascata: qualquer serie da coluna e a barra da categoria ("0:c"); duas series da mesma coluna = a mesma barra.
Public Sub Test_DeltaPick_Waterfall_ColumnIsTheBar()
    Dim wb As Workbook, ws As Worksheet, c As Chart, i As Long, s1 As Long, s2 As Long
    On Error GoTo Fail
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Receita", 100)
    ws.Range("A3:B3").Value = Array("Custo", -30)
    ws.Range("A4:B4").Value = Array("Lucro", "e")
    ws.Range("A1:B4").Select
    Core_Insert "waterfall"
    Set c = Core_ActiveChart()
    c.Parent.Activate
    Core_DeltaPick
    AssertTrue DeltaPick_IsArmedFor(c)
    For i = 1 To c.SeriesCollection.Count
        If Delta_IsColumnSeries(c.SeriesCollection(i)) Then
            If s1 = 0 Then
                s1 = i
            ElseIf s2 = 0 Then
                s2 = i
            End If
        End If
    Next
    AssertTrue s1 > 0, "serie de colunas na cascata"
    c.SeriesCollection(s1).Points(1).Select
    AssertEqual "0:1", DeltaPick_FromRef()
    If s2 > 0 Then
        c.SeriesCollection(s2).Points(1).Select
        AssertEqual DELTAPICK_MSG_SAME, StatusText(), "outra serie da mesma coluna = a mesma barra"
    End If
    c.SeriesCollection(s1).Points(3).Select
    AssertFalse DeltaPick_IsArmed()
    AssertTrue HasShape(c, "tqq_ann_delta_1")
    AssertEqual "kind=delta;v=1;from=0:1;to=0:3;fmt=both", c.Shapes("tqq_ann_delta_1").AlternativeText
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' Troca de planilha e fechamento da pasta com o modo armado: desarma, apaga a moldura, nada pendente.
Public Sub Test_DeltaPick_SheetSwitch_AndBookClose_Cancel()
    Dim wb As Workbook, c As Chart
    On Error GoTo Fail
    Watchers_Start
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(1).Select
    AssertTrue HasShape(c, DELTAPICK_MARK)
    wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
    AssertFalse DeltaPick_IsArmed(), "outra planilha ativada cancela"
    AssertFalse HasShape(c, DELTAPICK_MARK), "moldura removida"
    wb.Worksheets(1).Activate
    c.Parent.Activate
    Core_DeltaPick
    AssertTrue DeltaPick_IsArmedFor(c)
    wb.Close SaveChanges:=False
    Set wb = Nothing
    AssertFalse DeltaPick_IsArmed(), "fechar a pasta cancela"
    AssertEqual 0, DeltaPick_PendingTimers(), "o tempo de 60 s foi cancelado"
    AssertEqual 0, DeltaPick_TimerFailures()
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' O grafico armado foi apagado: Esc desarma sem erro e devolve a barra de status; uma moldura que sobrou (pasta salva com o modo armado)
' sai no proximo refresh, e a do grafico armado fica.
Public Sub Test_DeltaPick_DeletedChart_AndStaleMarkCleanup()
    Dim wb As Workbook, c As Chart, m As Shape
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(1).Select
    AssertTrue HasShape(c, DELTAPICK_MARK)
    Core_SafeRefresh c
    AssertTrue HasShape(c, DELTAPICK_MARK), "refresh do grafico armado mantem a moldura"
    c.Parent.Delete
    DeltaPick_EscKey
    AssertFalse DeltaPick_IsArmed()
    AssertEqual DELTAPICK_MSG_CANCEL, StatusText()
    DeltaPick_Disarm ""
    AssertEqual "", StatusText()
    Teardown wb
    Set wb = NewTempWorkbook()
    Fixture(wb).Select
    Core_Insert "clustered"
    Set c = Core_ActiveChart()
    Set m = c.Shapes.AddShape(msoShapeRectangle, 10, 10, 20, 20)
    m.Name = DELTAPICK_MARK
    Core_SafeRefresh c
    AssertFalse HasShape(c, DELTAPICK_MARK), "moldura de grafico nao armado apagada no refresh"
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' Pixels (pt x 4/3 a 100%) do centro do ponto p da serie s; False (e registra) se o GetChartElement deste Excel nao der esse ponto.
Private Function PointPixels(c As Chart, s As Long, p As Long, ByRef x As Long, ByRef y As Long) As Boolean
    Dim pt As Point, id As Long, a1 As Long, a2 As Long
    Set pt = c.SeriesCollection(s).Points(p)
    x = CLng((pt.Left + pt.Width / 2) * 4 / 3): y = CLng((pt.Top + pt.Height / 2) * 4 / 3)
    c.GetChartElement x, y, id, a1, a2
    PointPixels = (id = xlSeries And a1 = s And a2 = p)
    If Not PointPixels Then Errors_Log "TestDeltaPick: GetChartElement(" & x & "," & y & ") deu " & id & "/" & a1 & "/" & a2 & " para " & s & "/" & p & ": trecho do MouseUp pulado"
End Function

' O OnTime do prazo dispara: desarma mesmo com Now um pouco antes do prazo (Date exato podia nao vencer por um ulp e deixar o modo armado e
' o Esc preso; a folga e de 2 s, B6). Disparar de novo, ja desarmado, so devolve o Esc (sem erro).
Public Sub Test_DeltaPick_TimeoutFired_DisarmsWithoutDeadlineCompare()
    Dim wb As Workbook, c As Chart
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(1).Select
    AssertTrue HasShape(c, DELTAPICK_MARK)
    DeltaPick_FireTimeoutForTest                    ' o prazo de 60 s ainda nao venceu (Now < mDeadline)
    AssertFalse DeltaPick_IsArmed(), "OnTime disparado desarma sem olhar o relogio"
    AssertEqual DELTAPICK_MSG_TIMEOUT, StatusText()
    AssertFalse HasShape(c, DELTAPICK_MARK), "moldura removida"
    AssertEqual 1, DeltaPick_PendingTimers(), "so o aviso de 4 s"
    AssertEqual 0, DeltaPick_TimerFailures(), "nenhum cancelamento do OnTime ja disparado"
    DeltaPick_TimeoutFired
    AssertFalse DeltaPick_IsArmed()
    DeltaPick_Disarm ""
    AssertEqual 0, DeltaPick_PendingTimers()
    AssertEqual 0, DeltaPick_TimerFailures()
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' B6: um OnTime velho (de um modo anterior cujo cancelamento falhou, ou o de antes da renovacao do prazo) que dispara com o modo NOVO
' armado e o prazo ainda longe nao desarma, nem esquece o OnTime vivo; o disparo do OnTime do modo atual (ate 2 s antes do prazo) desarma.
Public Sub Test_DeltaPick_StaleTimeout_DoesNotDisarmNewArm()
    Dim wb As Workbook, c As Chart
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    DeltaPick_TimeoutFired                                   ' disparo velho: o prazo deste modo vence daqui a ~60 s
    AssertTrue DeltaPick_IsArmedFor(c), "temporizador velho nao desarma o modo novo"
    AssertEqual DELTAPICK_MSG_FROM, StatusText()
    AssertEqual 1, DeltaPick_PendingTimers(), "o OnTime vivo continua registrado (cancelavel no desarme)"
    DeltaPick_FireTimeoutForTest                             ' o disparo do OnTime do modo atual (1 s antes do prazo: tolerancia)
    AssertFalse DeltaPick_IsArmed(), "o OnTime do modo atual desarma"
    AssertEqual DELTAPICK_MSG_TIMEOUT, StatusText()
    DeltaPick_Disarm ""
    AssertEqual 0, DeltaPick_PendingTimers()
    AssertEqual 0, DeltaPick_TimerFailures()
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' B6: o prazo de 60 s conta do ultimo clique ACEITO: o clique de origem 40 s depois de armar deixa de novo ~60 s (OnTime cancelado e
' reagendado com o mesmo procedimento, sem empilhar; prazo e horario agendado coerentes). Clique que nao vira origem nao renova.
Public Sub Test_DeltaPick_OriginClick_RenewsDeadline()
    Dim wb As Workbook, c As Chart, s As Long
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    s = Delta_SeriesIndex(c, 1)
    AssertTrue DeltaPick_SecondsLeft() > 55 And DeltaPick_SecondsLeft() <= 61, "armado: ~60 s (" & DeltaPick_SecondsLeft() & ")"
    DeltaPick_AgeForTest 40
    AssertTrue DeltaPick_SecondsLeft() > 15 And DeltaPick_SecondsLeft() <= 21, "40 s simulados: ~20 s (" & DeltaPick_SecondsLeft() & ")"
    DeltaPick_OnSelect c, xlDataLabel, s, 1                  ' rotulo: nao e clique aceito
    AssertTrue DeltaPick_SecondsLeft() <= 21, "clique fora da barra nao renova (" & DeltaPick_SecondsLeft() & ")"
    c.SeriesCollection(s).Points(1).Select
    AssertEqual "1:1", DeltaPick_FromRef()
    AssertTrue DeltaPick_SecondsLeft() > 55 And DeltaPick_SecondsLeft() <= 61, "origem aceita: ~60 s de novo (" & DeltaPick_SecondsLeft() & ")"
    AssertTrue Abs(DeltaPick_TimerSecondsLeft() - DeltaPick_SecondsLeft()) < 1.5, "OnTime reagendado no novo prazo (" & DeltaPick_TimerSecondsLeft() & ")"
    AssertEqual 1, DeltaPick_PendingTimers(), "um so OnTime do prazo"
    AssertEqual 0, DeltaPick_TimerFailures(), "o OnTime antigo foi cancelado com o horario agendado"
    DeltaPick_Disarm ""
    AssertEqual 0, DeltaPick_PendingTimers()
    AssertEqual 0, DeltaPick_TimerFailures()
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' Um clique fisico chega como Select (ponto pelo Excel) e MouseUp (ponto pelo GetChartElement); se os dois discordam (borda, zoom, DPI)
' o MouseUp logo depois de um Select aceito e ignorado: um clique nunca completa o delta. Dois cliques rapidos em barras DIFERENTES pelo
' mesmo caminho (Select -> Select ou MouseUp -> MouseUp) continuam valendo.
Public Sub Test_DeltaPick_SelectThenMouseUp_IsOneClick()
    Dim wb As Workbook, c As Chart, s As Long, x As Long, y As Long, x2 As Long, y2 As Long
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    ForceLayout c
    s = Delta_SeriesIndex(c, 2)
    If PointPixels(c, s, 3, x, y) And PointPixels(c, s, 2, x2, y2) Then
        DeltaPick_OnSelect c, xlSeries, s, 1
        AssertEqual "2:1", DeltaPick_FromRef()
        DeltaPick_OnMouseUp c, x, y                     ' o mesmo clique, hit test noutra barra
        AssertEqual "2:1", DeltaPick_FromRef(), "MouseUp do mesmo clique nao vira destino"
        AssertTrue DeltaPick_IsArmed(), "segue armado"
        AssertFalse HasShape(c, "tqq_ann_delta_1"), "um clique nao insere delta"
        DeltaPick_Disarm ""
        ' ordem real Select -> MouseUp na mesma barra: o Select (evento real) escolhe, o MouseUp nao muda nada
        c.ChartArea.Select
        Core_DeltaPick
        c.SeriesCollection(s).Points(2).Select
        DeltaPick_OnMouseUp c, x2, y2
        AssertEqual "2:2", DeltaPick_FromRef()
        AssertEqual DELTAPICK_MSG_TO, StatusText(), "nada de 'mesma barra' no mesmo clique"
        ' Select -> Select rapido em outra barra: e o 2o clique (destino)
        DeltaPick_OnSelect c, xlSeries, s, 3
        AssertFalse DeltaPick_IsArmed(), "Select -> Select: destino"
        AssertEqual "kind=delta;v=1;from=2:2;to=2:3;fmt=both", c.Shapes("tqq_ann_delta_1").AlternativeText
        ' MouseUp -> MouseUp rapido em outra barra: tambem e o 2o clique
        ForceLayout c                                   ' a seta abriu sala: a area de plotagem mudou
        If PointPixels(c, s, 3, x, y) And PointPixels(c, s, 2, x2, y2) Then
            c.ChartArea.Select
            Core_DeltaPick
            AssertTrue DeltaPick_IsArmedFor(c)
            DeltaPick_OnMouseUp c, x2, y2
            DeltaPick_OnMouseUp c, x, y
            AssertFalse DeltaPick_IsArmed(), "MouseUp -> MouseUp: destino"
            AssertEqual "kind=delta;v=1;from=2:2;to=2:3;fmt=both", c.Shapes("tqq_ann_delta_2").AlternativeText
        End If
    End If
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' B6: a janela de 0,6 s do mesmo clique comeca DEPOIS de desenhar a moldura: num grafico pesado (desenho de 0,8 s) o MouseUp do clique que
' escolheu a origem pelo Select ainda e o mesmo clique (hit test noutra barra) e nao completa o delta.
Public Sub Test_DeltaPick_SlowMark_SameClickWindowStartsAfterDraw()
    Dim wb As Workbook, c As Chart, s As Long, x As Long, y As Long
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    ForceLayout c
    s = Delta_SeriesIndex(c, 2)
    If PointPixels(c, s, 3, x, y) Then
        DeltaPick_SlowMarkForTest 0.8
        DeltaPick_OnSelect c, xlSeries, s, 1
        AssertEqual "2:1", DeltaPick_FromRef()
        AssertTrue HasShape(c, DELTAPICK_MARK), "moldura desenhada"
        DeltaPick_OnMouseUp c, x, y                     ' o MouseUp do mesmo clique, logo depois do desenho lento
        AssertTrue DeltaPick_IsArmed(), "o MouseUp do mesmo clique nao vira destino"
        AssertEqual "2:1", DeltaPick_FromRef()
        AssertFalse HasShape(c, "tqq_ann_delta_1"), "um clique nao insere delta"
    End If
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' Rotulo de dados, linhas de serie, formas (setas, a propria moldura): "clique numa barra", segue armado; area/eixo/legenda: silencio.
Public Sub Test_DeltaPick_NonBarElements_StayArmed()
    Dim wb As Workbook, c As Chart, s As Long
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    s = Delta_SeriesIndex(c, 1)
    DeltaPick_OnSelect c, xlChartArea, 0, 0
    AssertEqual DELTAPICK_MSG_FROM, StatusText(), "area do grafico: sem mensagem"
    DeltaPick_OnSelect c, xlDataLabel, s, 1
    AssertEqual DELTAPICK_MSG_NOTBAR, StatusText(), "rotulo de dados"
    AssertTrue DeltaPick_IsArmed()
    AssertEqual "", DeltaPick_FromRef()
    Application.StatusBar = DELTAPICK_MSG_FROM
    DeltaPick_OnSelect c, xlShape, 0, 0
    AssertEqual DELTAPICK_MSG_NOTBAR, StatusText(), "forma (seta/moldura)"
    Application.StatusBar = DELTAPICK_MSG_FROM
    DeltaPick_OnSelect c, xlSeriesLines, s, 0
    AssertEqual DELTAPICK_MSG_NOTBAR, StatusText(), "linhas de serie"
    AssertTrue DeltaPick_IsArmed()
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' Esc com o estado do VBA perdido (desarmado mas a tecla ainda ligada): devolve o Esc ao Excel sem erro e sem mexer na barra de status.
Public Sub Test_DeltaPick_EscWhenNotArmed_SelfHeals()
    Dim wb As Workbook
    On Error GoTo Fail
    DeltaPick_Disarm ""
    Application.StatusBar = False
    DeltaPick_EscKey
    AssertFalse DeltaPick_IsArmed()
    AssertEqual "", StatusText()
    AssertEqual 0, DeltaPick_PendingTimers()
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' Core_OnUnload (core fechando) desarma: moldura, Esc e nenhum OnTime apontando para o core fechado.
Public Sub Test_DeltaPick_CoreUnload_Disarms()
    Dim wb As Workbook, c As Chart
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Set c = ArmedChart(wb, "clustered")
    c.SeriesCollection(Delta_SeriesIndex(c, 1)).Points(1).Select
    AssertTrue HasShape(c, DELTAPICK_MARK)
    Core_OnUnload
    AssertFalse DeltaPick_IsArmed()
    AssertFalse HasShape(c, DELTAPICK_MARK)
    AssertEqual 0, DeltaPick_PendingTimers()
    AssertEqual 0, DeltaPick_TimerFailures()
    AssertEqual "", StatusText()
    Teardown wb                                     ' Core_OnUnload ja desligou os observadores (como o Teardown de todo teste)
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' DeltaPick_Arm falha (grafico apagado entre a escolha e o armar): desarma e avisa o usuario.
Public Sub Test_DeltaPick_ArmFailure_Notifies()
    Dim wb As Workbook, c As Chart
    On Error GoTo Fail
    Set wb = NewTempWorkbook()
    Fixture(wb).Select
    Core_Insert "clustered"
    Set c = Core_ActiveChart()
    c.Parent.Delete
    Core_ClearLastNotice
    DeltaPick_Arm c
    AssertFalse DeltaPick_IsArmed()
    AssertEqual DELTAPICK_MSG_ARMFAIL, Core_LastNotice()
    AssertEqual 0, DeltaPick_PendingTimers()
    AssertEqual "", StatusText()
    Teardown wb
    Exit Sub
Fail:
    FailTeardown wb
End Sub

' B4: "Delta: valor / % / ambos" alterna o texto de TODAS as setas de delta (nenhuma selecionada) ou so da selecionada (rotulo, perna ou
' linha dela); sem delta no grafico, avisa.
Public Sub Test_Delta_CycleFormat_AllOrSelected()
    Dim wb As Workbook, rng As Range, c As Chart
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "clustered"
    Set c = Core_ActiveChart()
    Core_AddDelta "1:1", "1:3"
    Core_AddDelta "2:1", "2:2"
    c.Parent.Activate
    c.ChartArea.Select
    Core_DeltaCycleFormat
    AssertEqual "abs", Ann_ParseParam(c.Shapes("tqq_ann_delta_1").AlternativeText, "fmt"), "sem seta selecionada: todas"
    AssertEqual "abs", Ann_ParseParam(c.Shapes("tqq_ann_delta_2").AlternativeText, "fmt")
    AssertEqual Delta_Text(10, 30, "abs", Style_NumberFormat(Registry_GetOption(c, "numfmt", "int"))), CStr(c.Shapes("tqq_ann_delta_1_lbl").TextFrame2.TextRange.Text)
    c.Shapes("tqq_ann_delta_2_lbl").Select
    AssertEqual "tqq_ann_delta_2", Ann_SelectedDeltaName()
    Core_DeltaCycleFormat
    AssertEqual "abs", Ann_ParseParam(c.Shapes("tqq_ann_delta_1").AlternativeText, "fmt"), "so a selecionada muda"
    AssertEqual "pct", Ann_ParseParam(c.Shapes("tqq_ann_delta_2").AlternativeText, "fmt")
    Ann_RemoveAll c
    c.ChartArea.Select
    Core_ClearLastNotice
    Core_DeltaCycleFormat
    AssertEqual DELTA_NONE, Core_LastNotice(), "sem delta: aviso"
    Teardown wb
End Sub

' B4: a seta selecionada pode ser a perna (_t1/_t2) ou a linha; qualquer outra selecao (area do grafico) devolve "".
Public Sub Test_Delta_SelectedName_LegsAndLine()
    Dim wb As Workbook, c As Chart
    Set wb = NewTempWorkbook()
    Fixture(wb).Select
    Core_Insert "clustered"
    Set c = Core_ActiveChart()
    Core_AddDelta "1:1", "1:3"
    c.Parent.Activate
    c.Shapes("tqq_ann_delta_1_t1").Select
    AssertEqual "tqq_ann_delta_1", Ann_SelectedDeltaName(), "perna da origem"
    c.Shapes("tqq_ann_delta_1_t2").Select
    AssertEqual "tqq_ann_delta_1", Ann_SelectedDeltaName(), "perna do destino"
    c.Shapes("tqq_ann_delta_1").Select
    AssertEqual "tqq_ann_delta_1", Ann_SelectedDeltaName(), "linha"
    c.ChartArea.Select
    AssertEqual "", Ann_SelectedDeltaName(), "area do grafico"
    Teardown wb
End Sub
