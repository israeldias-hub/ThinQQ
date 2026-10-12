Attribute VB_Name = "TestPowerPoint"
Option Explicit

Private Sub Teardown(wb As Workbook)
    Watchers_Stop
    CloseTemp wb
End Sub

Private Function Fixture(wb As Workbook) As Range
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2024, 10, 5): ws.Range("A3:C3").Value = Array(2025, 20, 6)
    Set Fixture = ws.Range("A1:C3")
End Function

' Processos POWERPNT.EXE (consulta WMI; o moniker winmgmts leva duas barras invertidas, montadas com Chr(92)).
Private Function PowerPointProcessCount() As Long
    PowerPointProcessCount = GetObject("winmgmts:" & Chr(92) & Chr(92) & "." & Chr(92) & "root" & Chr(92) & "cimv2").ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'POWERPNT.EXE'").Count
End Function

' False se a area de transferencia do Windows nao pode ser aberta agora (outro programa a mantem aberta): a colagem real nao tem como rodar.
Private Function ClipboardUsable(cht As Chart) As Boolean
    On Error Resume Next
    cht.CopyPicture xlScreen, xlPicture
    ClipboardUsable = (Err.Number = 0)
    Err.Clear
End Function

Public Sub Test_Ppt_FitSize()
    Dim w As Double, h As Double
    Ppt_FitSize 460, 250, 960, 540, w, h
    AssertAlmost 768, w, 0.001: AssertAlmost 250# * 768 / 460, h, 0.001, "80% da largura, proporcao mantida"
    Ppt_FitSize 460, 280, 960, 540, w, h
    AssertAlmost 432, h, 0.001, "o grafico padrao (460 x 280) e limitado pela altura": AssertAlmost 460# * 432 / 280, w, 0.001
    Ppt_FitSize 200, 600, 960, 540, w, h
    AssertAlmost 432, h, 0.001, "limitado por 80% da altura": AssertAlmost 200# * 432 / 600, w, 0.001
    Ppt_FitSize 0, 0, 960, 540, w, h
    AssertEqual 0, w, "tamanho invalido fica como esta"
End Sub

' Review Focus 5 (parte 1): pasta nao salva e vinculo.
Public Sub Test_Ppt_LinkedPrecheck_UnsavedWorkbook()
    Dim wb As Workbook, p As String
    Set wb = NewTempWorkbook()
    AssertEqual PPT_ERR_UNSAVED, Ppt_LinkedPrecheck(wb)
    p = TempDir() & "\ppt.xlsx"
    wb.SaveAs p, xlOpenXMLWorkbook
    AssertEqual "", Ppt_LinkedPrecheck(wb), "pasta salva"
    CloseTemp wb
    AssertEqual PPT_ERR_NOT_INSTALLED, Ppt_ErrorText(429, "x")
    AssertTrue InStr(Ppt_ErrorText(70, "Permissão negada"), "Permissão negada") > 0
    AssertEqual PPT_ERR_CLIPBOARD, Ppt_ErrorText(-2147221040, "x"), "CLIPBRD_E_CANT_OPEN: mensagem propria"
    AssertTrue Ppt_IsKnownFailure(429) And Ppt_IsKnownFailure(-2147221040) And Not Ppt_IsKnownFailure(70), "falhas esperadas viram aviso proprio"
End Sub

Public Sub Test_Core_CopyToPowerPoint_Unsaved_Linked_NoticeBeforePowerPoint()
    Dim wb As Workbook, rng As Range
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    rng.Select
    Core_Insert "stacked"
    Core_ClearLastNotice
    Core_CopyToPowerPoint "chart"
    AssertEqual PPT_ERR_UNSAVED, Core_LastNotice(), "vinculado com pasta nao salva: so o aviso"
    Core_ClearLastNotice
    Core_CopyToPowerPoint "xyz"
    AssertEqual "", Core_LastNotice(), "modo desconhecido: nada (so log)"
    rng.Worksheet.Range("A1").Select
    Core_ClearLastNotice
    Core_CopyToPowerPoint "image"
    AssertEqual "", Core_LastNotice(), "sem grafico: nada"
    Teardown wb
End Sub

' Review Focus 5 (parte 2): colagem real. Roda SO quando nenhum PowerPoint esta aberto (aplicativo de instancia unica: CreateObject anexaria
' ao do usuario) e o PowerPoint existe; senao sai verde registrando o motivo. Apresentacao sem janela; a instancia e encerrada ao final.
Public Sub Test_Ppt_RealPaste_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, c As Chart, en As Long, ed As String, shp As Object
    If Ppt_IsRunning() Then Errors_Log "TestPowerPoint: pulado (ha um PowerPoint aberto)": Exit Sub
    ' a ROT (GetObject) nao ve um PowerPoint aberto so para incorporacao: confere tambem os processos (regra global: nenhum POWERPNT)
    If PowerPointProcessCount() > 0 Then Errors_Log "TestPowerPoint: pulado (processo POWERPNT existente)": Exit Sub
    On Error Resume Next
    Set app = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    If app Is Nothing Then Errors_Log "TestPowerPoint: pulado (PowerPoint nao instalado)": Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0)
    Set slide = pres.Slides.Add(1, 12)
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\ppt-src.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ann_AddCagr c, 1, 2
    If Not ClipboardUsable(c) Then Errors_Log "TestPowerPoint: colagem real pulada (area de transferencia do Windows indisponivel)": GoTo Cleanup
    Ppt_SetAppForTest app
    Core_ClearLastNotice
    Core_CopyToPowerPoint "image"
    AssertEqual 1, slide.Shapes.Count, "imagem colada"
    AssertEqual 13, slide.Shapes(1).Type, "msoPicture"
    AssertAlmost 432, CDbl(slide.Shapes(1).Height), 2, "limitado pela altura (432 pt = 80% de 540): largura ~710 (" & slide.Shapes(1).Width & ")"
    AssertTrue InStr(Core_LastNotice(), "slide 1") > 0 And InStr(Core_LastNotice(), "como imagem") > 0, Core_LastNotice()
    Core_ClearLastNotice
    Core_CopyToPowerPoint "chart"
    AssertEqual 2, slide.Shapes.Count, "grafico colado"
    Set shp = slide.Shapes(2)
    AssertEqual -1, shp.HasChart, "gráfico do PowerPoint"
    AssertEqual -1, CLng(shp.Chart.ChartData.IsLinked), "vinculado a pasta salva"
    AssertTrue shp.Chart.Shapes.Count >= 2, "formas da seta CAGR preservadas: " & shp.Chart.Shapes.Count
    AssertTrue InStr(Core_LastNotice(), "gráfico vinculado") > 0, Core_LastNotice()
    rng.Worksheet.Range("B2").Value = 11
    Core_ClearLastNotice
    Core_CopyToPowerPoint "chart"
    AssertTrue InStr(Core_LastNotice(), PPT_NOTE_DIRTY) > 0, "aviso de alteracoes nao salvas"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Ppt_SetAppForTest Nothing
    If Not pres Is Nothing Then pres.Close
    If Not app Is Nothing Then If app.Presentations.Count = 0 Then app.Quit      ' alguem anexou a esta instancia e abriu algo: nao encerra
    Set app = Nothing
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPowerPoint", ed
End Sub

' Mesmas regras de seguranca do teste anterior (so cria o PowerPoint se nenhum estiver aberto). Devolve Nothing (e registra o motivo) se pulou.
Private Function StartScratchPowerPoint() As Object
    Dim app As Object
    If Ppt_IsRunning() Then Errors_Log "TestPowerPoint: pulado (ha um PowerPoint aberto)": Exit Function
    If PowerPointProcessCount() > 0 Then Errors_Log "TestPowerPoint: pulado (processo POWERPNT existente)": Exit Function
    On Error Resume Next
    Set app = CreateObject("PowerPoint.Application")
    On Error GoTo 0
    If app Is Nothing Then Errors_Log "TestPowerPoint: pulado (PowerPoint nao instalado)": Exit Function
    Set StartScratchPowerPoint = app
End Function

' Encerra o PowerPoint criado pelo teste: fecha as apresentacoes dele (so as dele: a instancia e nossa) e da Quit.
Private Sub StopScratchPowerPoint(app As Object)
    Dim n As Long
    On Error Resume Next
    Ppt_SetFailForTest ""
    Ppt_SetAppForTest Nothing
    If app Is Nothing Then Exit Sub
    n = app.Presentations.Count
    Do While n > 0
        app.Presentations(n).Saved = -1
        app.Presentations(n).Close
        n = n - 1
    Loop
    If app.Presentations.Count = 0 Then app.Quit
End Sub

' Sem apresentacao aberta (gancho "slide"): a apresentacao criada pela chamada que falhou e fechada; o gancho nao vaza para a chamada
' seguinte (nem numa saida antecipada); sem falha, a chamada cria a apresentacao e cola no slide 1. Nao usa a area de transferencia ate o gancho.
Public Sub Test_Ppt_NoPresentation_FailureClosesWhatItCreated_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, en As Long, ed As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\ppt-src3.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Ppt_SetAppForTest app
    AssertEqual 0, app.Presentations.Count, "ponto de partida: sem apresentacao"
    Core_ClearLastNotice
    Ppt_SetFailForTest "slide"
    Core_CopyToPowerPoint "image"
    AssertEqual 0, app.Presentations.Count, "a apresentacao criada pela chamada que falhou foi fechada"
    AssertTrue InStr(Core_LastNotice(), "falha injetada (slide)") > 0, Core_LastNotice()
    Ppt_SetFailForTest "slide"
    Core_CopyToPowerPoint "xyz"                      ' saida antecipada: o gancho nao pode ficar armado
    Core_ClearLastNotice
    Core_CopyToPowerPoint "image"
    If Core_LastNotice() = PPT_ERR_CLIPBOARD Then
        Errors_Log "TestPowerPoint: parte da colagem pulada (area de transferencia do Windows indisponivel)"
        AssertEqual 0, app.Presentations.Count, "sem area de transferencia: a apresentacao criada e fechada"
    Else
        AssertEqual 1, app.Presentations.Count, "apresentacao criada"
        AssertEqual 1, app.Presentations(1).Slides(1).Shapes.Count, "imagem no slide novo"
        AssertTrue InStr(Core_LastNotice(), "slide 1 de '" & app.Presentations(1).Name & "'") > 0, Core_LastNotice()
    End If
Cleanup:
    en = Err.Number: ed = Err.Description
    StopScratchPowerPoint app
    Set app = Nothing
    On Error Resume Next
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPowerPoint", ed
End Sub

' Falha no ajuste (gancho "place"): a forma colada sai do slide e o aviso explica; o gancho nao vaza; a copia do Excel (CutCopyMode) e
' cancelada. Com a area de transferencia indisponivel o teste valida so o aviso proprio e que nada foi colado.
Public Sub Test_Ppt_FailureAfterPaste_RemovesPastedShape_WhenAvailable()
    Dim wb As Workbook, rng As Range, app As Object, pres As Object, slide As Object, c As Chart, en As Long, ed As String
    Set app = StartScratchPowerPoint()
    If app Is Nothing Then Exit Sub
    On Error GoTo Cleanup
    Set pres = app.Presentations.Add(0)
    Set slide = pres.Slides.Add(1, 12)
    Set wb = NewTempWorkbook(): Set rng = Fixture(wb)
    wb.SaveAs TempDir() & "\ppt-src2.xlsx", xlOpenXMLWorkbook
    rng.Select
    Core_Insert "stacked"
    Set c = Core_ActiveChart()
    Ppt_SetAppForTest app
    If Not ClipboardUsable(c) Then
        Errors_Log "TestPowerPoint: colagem real pulada (area de transferencia do Windows indisponivel); validando so o aviso"
        Core_ClearLastNotice
        Core_CopyToPowerPoint "image"
        AssertEqual 0, slide.Shapes.Count, "nada colado"
        AssertEqual PPT_ERR_CLIPBOARD, Core_LastNotice()
        GoTo Cleanup
    End If
    Core_ClearLastNotice
    Ppt_SetFailForTest "place"
    Core_CopyToPowerPoint "image"
    AssertEqual 0, slide.Shapes.Count, "imagem colada e removida quando o ajuste falha"
    AssertTrue InStr(Core_LastNotice(), "Não foi possível colar no PowerPoint") > 0 And InStr(Core_LastNotice(), "falha injetada (place)") > 0, Core_LastNotice()
    Core_ClearLastNotice
    Ppt_SetFailForTest "place"
    Core_CopyToPowerPoint "chart"
    AssertEqual 0, slide.Shapes.Count, "grafico colado e removido quando o ajuste falha"
    AssertEqual 0, CLng(Application.CutCopyMode), "modo de copia do Excel cancelado na falha"
    Ppt_SetFailForTest "place"
    Core_CopyToPowerPoint "xyz"                      ' saida antecipada: o gancho nao pode ficar armado
    Core_ClearLastNotice
    Core_CopyToPowerPoint "image"
    AssertEqual 1, slide.Shapes.Count, "gancho desarmado: a chamada seguinte cola"
    AssertTrue InStr(Core_LastNotice(), "como imagem") > 0, Core_LastNotice()
    Core_ClearLastNotice
    Core_CopyToPowerPoint "chart"
    AssertEqual 2, slide.Shapes.Count, "grafico vinculado colado"
    AssertEqual 0, CLng(Application.CutCopyMode), "modo de copia do Excel cancelado apos colar o grafico"
Cleanup:
    en = Err.Number: ed = Err.Description
    StopScratchPowerPoint app
    Set app = Nothing
    On Error Resume Next
    If Not wb Is Nothing Then Teardown wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestPowerPoint", ed
End Sub
