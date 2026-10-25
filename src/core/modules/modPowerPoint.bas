Attribute VB_Name = "modPowerPoint"
Option Explicit
' Copiar o grafico para o PowerPoint por automacao tardia (CreateObject/GetObject; sem referencia, sem Declare; constantes numericas).
' Imagem: Chart.CopyPicture (xlScreen, xlPicture) + Shapes.Paste (funciona com o Excel invisivel: medido). Grafico vinculado: ChartObject.Copy
' + Shapes.Paste cria um grafico do PowerPoint com ChartData.IsLinked = True e as formas das anotacoes preservadas (medido); exige a pasta
' SALVA (com pasta nova o vinculo aponta para um arquivo inexistente). O PowerPoint e aplicativo de instancia unica: GetObject anexa ao
' aberto; CreateObject abre um novo. Application.Visible = msoFalse e recusado pelo PowerPoint (os testes usam Presentations.Add(msoFalse)).
Private mTestApp As Object
Private mTestFail As String                     ' gancho de teste: ponto de falha injetada ("place", "slide"); so no build Test
Private mCreatedApp As Boolean                  ' a ultima Ppt_GetApp abriu o PowerPoint (nao havia um aberto): quem abre desfaz na falha
Private mTestPin As Object                      ' gancho de teste: slide fixo de destino (o do proprio teste); so no build Test
Private mTestForceNew As Boolean                ' gancho de teste: sempre uma apresentacao nova (nunca janela nem Presentations(1))
Private mTestLastNew As Object                  ' gancho de teste: a ultima apresentacao criada por Ppt_TargetSlide (e o nome dela)
Private mTestLastNewName As String
Private mOpDepth As Long                        ' operacoes do core com o PowerPoint em andamento (Ppt_BeginOp/Ppt_EndOp)
Private mOpSince As Date                        ' quando a operacao mais externa comecou
Private Const PPT_OP_STALE_S As Long = 120      ' marca de operacao mais velha que isto e considerada esquecida
Private Const PPT_LAYOUT_BLANK As Long = 12     ' ppLayoutBlank
Private Const PPT_FIT As Double = 0.8           ' fracao do slide que o grafico ocupa (largura ou altura, a que limitar)
Private Const PPT_CLIPBRD_E_CANT_OPEN As Long = -2147221040     ' 0x800401D0: a area de transferencia esta aberta por outro programa
Private Const PPT_CLIP_TRIES As Long = 3        ' tentativas de copiar quando a area de transferencia esta ocupada
Private Const PPT_CLIP_WAIT As Single = 0.25    ' segundos entre as tentativas
Private Const PPT_RPC_E_CALL_REJECTED As Long = -2147418111     ' 0x80010001: o PowerPoint recusou a chamada (caixa de dialogo aberta, texto em edicao)
Private Const PPT_RPC_E_RETRYLATER As Long = -2147417846        ' 0x8001010A: o PowerPoint esta ocupado (mesma causa)
Private Const PPT_DEFAULT_SLIDE_W As Double = 960   ' slide 16:9 padrao (pt): usado quando o PowerPoint devolve um tamanho invalido
Private Const PPT_DEFAULT_SLIDE_H As Double = 540
Public Const PPT_ERR_NOT_INSTALLED As String = "O PowerPoint não está instalado neste computador (ou não pôde ser iniciado)."
Public Const PPT_ERR_UNSAVED As String = "Salve a pasta de trabalho antes de copiar o gráfico vinculado: o vínculo aponta para o arquivo salvo."
Public Const PPT_ERR_CLIPBOARD As String = "A área de transferência do Windows não está disponível (outro programa a mantém aberta). Tente novamente em instantes."
Public Const PPT_ERR_BUSY As String = "O PowerPoint está ocupado (uma caixa de diálogo aberta ou um texto em edição). Termine o que está fazendo nele e tente de novo."
Public Const PPT_NOTE_DIRTY As String = "A pasta tem alterações não salvas: enquanto ela estiver aberta no Excel o vínculo lê esses valores; fechada, o PowerPoint lê o arquivo salvo."

' Gancho de teste: instancia injetada (inerte fora do build Test).
Public Sub Ppt_SetAppForTest(app As Object)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    Set mTestApp = app
End Sub

' Gancho de teste: faz o proximo ponto de falha nomeado levantar erro (uso unico; "" desarma). Inerte fora do build Test.
Public Sub Ppt_SetFailForTest(point As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestFail = point
End Sub

Private Sub Ppt_Hook(point As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    If mTestFail <> point Then Exit Sub
    mTestFail = ""
    Err.Raise 5, "modPowerPoint", "falha injetada (" & point & ")"
End Sub

' Gancho de teste: fixa o destino de Ppt_TargetSlide no slide do proprio teste (Nothing desfixa) ou, com forceNew, faz cada chamada criar
' a sua apresentacao (sem janela no build Test). Assim um PowerPoint ou .pptx que o usuario abra na instancia do teste (o PowerPoint e de
' instancia unica) nunca vira destino. Inerte fora do build Test.
Public Sub Ppt_PinTargetForTest(slide As Object, Optional forceNew As Boolean = False)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    Set mTestPin = slide
    mTestForceNew = forceNew
End Sub

' Gancho de teste: a ultima apresentacao que Ppt_TargetSlide criou (Nothing se nenhuma) e o nome lido ao criar (vale depois de fechada).
Public Function Ppt_LastCreatedPresForTest() As Object
    If BUILD_CONFIG <> "Test" Then Exit Function
    Set Ppt_LastCreatedPresForTest = mTestLastNew
End Function

Public Function Ppt_LastCreatedPresNameForTest() As String
    If BUILD_CONFIG <> "Test" Then Exit Function
    Ppt_LastCreatedPresNameForTest = mTestLastNewName
End Function

' Gancho de teste: desarma todos os ganchos do modulo (instancia injetada, falha, destino fixo, ultima apresentacao criada).
Public Sub Ppt_ResetTestHooks()
    Set mTestApp = Nothing
    mTestFail = ""
    Set mTestPin = Nothing
    mTestForceNew = False
    Set mTestLastNew = Nothing
    mTestLastNewName = ""
End Sub

' Ponto de falha injetada para quem chama (Core_CopyToPowerPoint: "after", depois de colar). Inerte fora do build Test.
Public Sub Ppt_TestHook(point As String)
    Ppt_Hook point
End Sub

Public Function Ppt_IsRunning() As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = GetObject(, "PowerPoint.Application")
    Ppt_IsRunning = (Err.Number = 0) And (Not o Is Nothing)
    Err.Clear
End Function

' Processos POWERPNT.EXE (WMI, ligacao tardia; o moniker leva duas barras invertidas, montadas com Chr(92)); -1 se a consulta falha.
Public Function Ppt_ProcessCount() As Long
    On Error GoTo Fail
    Ppt_ProcessCount = GetObject("winmgmts:" & Chr(92) & Chr(92) & "." & Chr(92) & "root" & Chr(92) & "cimv2").ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'POWERPNT.EXE'").Count
    Exit Function
Fail:
    Ppt_ProcessCount = -1
End Function

' O PowerPoint que CreateObject devolveu e nosso (encerrado numa falha) so se nao havia NENHUM processo POWERPNT antes: CreateObject anexa a
' um PowerPoint aberto so para incorporacao (sem apresentacoes, fora da ROT, invisivel a GetObject). Contagem desconhecida (-1): nao e nosso.
Public Function Ppt_CreatedByUs(procsBefore As Long, app As Object) As Boolean
    If app Is Nothing Then Exit Function
    Ppt_CreatedByUs = (procsBefore = 0)
End Function

' Operacao do core com o PowerPoint em andamento (colagem, sincronizacao pedida pelo usuario ou pelo .ppam): o ciclo automatico
' (PptAuto_Tick) espera. Contador (as entradas se aninham: Core_PptSyncOpen chama Core_PptSyncPresentation); quem abre fecha em todo caminho.
Public Sub Ppt_BeginOp()
    If mOpDepth = 0 Then mOpSince = Now
    mOpDepth = mOpDepth + 1
End Sub

Public Sub Ppt_EndOp()
    If mOpDepth > 0 Then mOpDepth = mOpDepth - 1
End Sub

' True durante uma operacao do core com o PowerPoint. Uma marca esquecida (aberta ha mais de PPT_OP_STALE_S) e descartada e registrada,
' para o ciclo automatico nunca ficar adiado para sempre. Nunca propaga erro.
Public Function Ppt_InOperation() As Boolean
    On Error Resume Next
    If mOpDepth = 0 Then Exit Function
    If Now - mOpSince > TimeSerial(0, 0, PPT_OP_STALE_S) Then
        Errors_Log "Ppt_InOperation: marca de operacao esquecida descartada (" & mOpDepth & ")"
        mOpDepth = 0
        Exit Function
    End If
    Ppt_InOperation = True
End Function

' create = False: so um PowerPoint ja aberto (GetObject), sem mexer no indicador de "aberto por nos" (Ppt_AppWasCreated) de uma colagem em
' andamento; GetObject recusado pelo PowerPoint ocupado (caixa de dialogo aberta: Ppt_IsBusyError) LEVANTA esse erro (quem chama trata o
' ocupado: o ciclo automatico tenta de novo, o botao do Excel avisa), em vez de parecer "sem PowerPoint"; outro erro = Nothing. create =
' True: abre um se nao houver e registra se foi aberto por esta chamada (nunca levanta).
Public Function Ppt_GetApp(create As Boolean) As Object
    Dim o As Object, procs As Long, en As Long, ed As String
    If create Then mCreatedApp = False
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then
        PptSync_TestPoint "getapp"                 ' gancho: GetObject recusado (o numero do erro vem do teste)
        If Err.Number = 0 Then Set o = mTestApp     ' testes: so a instancia injetada (sem ela, Nothing): nunca o PowerPoint do usuario
    Else
        Set o = GetObject(, "PowerPoint.Application")
    End If
    en = Err.Number: ed = Err.Description
    Err.Clear
    On Error GoTo 0
    If o Is Nothing And Not create And Ppt_IsBusyError(en) Then Err.Raise en, "Ppt_GetApp", ed
    If BUILD_CONFIG = "Test" Then Set Ppt_GetApp = o: Exit Function
    On Error Resume Next
    If o Is Nothing And create Then
        procs = Ppt_ProcessCount()
        Set o = CreateObject("PowerPoint.Application")
        mCreatedApp = Ppt_CreatedByUs(procs, o)
    End If
    Err.Clear
    On Error GoTo 0
    Set Ppt_GetApp = o
End Function

' True se a ultima Ppt_GetApp abriu um PowerPoint novo (nunca no build Test: a instancia injetada nao e nossa para encerrar).
Public Function Ppt_AppWasCreated() As Boolean
    Ppt_AppWasCreated = mCreatedApp
End Function

' o se for um slide (tem SlideIndex: o mestre e os layouts que View.Slide devolve nos modos de mestre nao tem), senao Nothing.
Private Function AsSlide(o As Object) As Object
    Dim k As Long
    If o Is Nothing Then Exit Function
    On Error Resume Next
    k = o.SlideIndex
    If Err.Number = 0 And k > 0 Then Set AsSlide = o
    Err.Clear
End Function

' Slide de destino. Com janela: o da janela ativa (modo Normal), senao o selecionado nela (Classificador de slides, Estrutura de topicos),
' senao o ultimo slide da apresentacao DA JANELA (modos de mestre/anotacoes; sem janela ativa, a primeira janela): nunca Presentations(1),
' que pode ser outra apresentacao do usuario. Sem janela: o ultimo slide de Presentations(1); sem apresentacao (ou janela sem apresentacao
' legivel), uma apresentacao nova, devolvida em createdPres (quem chama a fecha se falhar depois). Um slide que esta chamada acrescentou a
' uma apresentacao ja aberta (sem slides) volta em addedSlide (quem chama o apaga se falhar depois). Propaga erros do PowerPoint.
Public Function Ppt_TargetSlide(app As Object, Optional ByRef createdPres As Object, Optional ByRef addedSlide As Object) As Object
    Dim pres As Object, slide As Object, win As Object, hasWin As Boolean, useOpen As Boolean
    Set createdPres = Nothing: Set addedSlide = Nothing
    useOpen = True
    If BUILD_CONFIG = "Test" Then
        If Not mTestPin Is Nothing Then Set Ppt_TargetSlide = mTestPin: Exit Function
        useOpen = Not mTestForceNew
    End If
    If useOpen Then hasWin = (app.Windows.Count > 0)
    If hasWin Then
        On Error Resume Next
        Set win = app.ActiveWindow
        If win Is Nothing Then Set win = app.Windows(1)
        If Not win Is Nothing Then
            Set slide = AsSlide(win.View.Slide)
            If slide Is Nothing Then Set slide = AsSlide(win.Selection.SlideRange(1))
            If slide Is Nothing Then Set pres = win.Presentation
        End If
        Err.Clear
        On Error GoTo 0
        If Not slide Is Nothing Then Set Ppt_TargetSlide = slide: Exit Function
    ElseIf useOpen Then
        If app.Presentations.Count > 0 Then Set pres = app.Presentations(1)
    End If
    If Not pres Is Nothing Then
        If pres.Slides.Count = 0 Then
            Set slide = pres.Slides.Add(1, PPT_LAYOUT_BLANK)
            Set addedSlide = slide
        Else
            Set slide = pres.Slides(pres.Slides.Count)
        End If
    Else
        ' msoTrue: com janela, para o usuario ver (no build Test sem janela: nao aparece na area de trabalho)
        Set pres = app.Presentations.Add(IIf(BUILD_CONFIG = "Test", 0, -1))
        Set createdPres = pres
        If BUILD_CONFIG = "Test" Then
            Set mTestLastNew = pres
            mTestLastNewName = CStr(pres.Name)
        End If
        Ppt_Hook "slide"
        Set slide = pres.Slides.Add(1, PPT_LAYOUT_BLANK)
    End If
    Set Ppt_TargetSlide = slide
End Function

' Tamanho do slide utilizavel: o 16:9 padrao (960 x 540) se o PowerPoint devolveu largura ou altura <= 0.
Private Sub SafeSlideSize(ByRef sw As Double, ByRef sh As Double)
    If sw <= 0 Or sh <= 0 Then sw = PPT_DEFAULT_SLIDE_W: sh = PPT_DEFAULT_SLIDE_H
End Sub

' Tamanho que cabe em PPT_FIT do slide mantendo a proporcao (pura). Nunca devolve tamanho <= 0: slide invalido vale como o 16:9 padrao e
' grafico de tamanho invalido ocupa a caixa inteira (PPT_FIT da largura e da altura do slide).
Public Sub Ppt_FitSize(w As Double, h As Double, slideW As Double, slideH As Double, ByRef outW As Double, ByRef outH As Double)
    Dim f As Double, sw As Double, sh As Double
    sw = slideW: sh = slideH
    SafeSlideSize sw, sh
    If w <= 0 Or h <= 0 Then outW = PPT_FIT * sw: outH = PPT_FIT * sh: Exit Sub
    f = PPT_FIT * sw / w
    If h * f > PPT_FIT * sh Then f = PPT_FIT * sh / h
    outW = w * f: outH = h * f
End Sub

' True se a colagem conta como feita para a moldura: imagem sempre; vinculado so se o PowerPoint manteve (1) ou nao deu para ler (-1). Nunca propaga erro.
Public Function Ppt_PasteKept(linked As Boolean, linkState As Long) As Boolean
    Ppt_PasteKept = (Not linked) Or linkState <> 0
End Function

' A colagem pede a moldura do intervalo de origem (1.13.0, X2b)? So o grafico VINCULADO que manteve o vinculo (Ppt_PasteKept); imagem nunca.
Public Function Ppt_FrameAfterCopy(linked As Boolean, linkState As Long) As Boolean
    Ppt_FrameAfterCopy = linked And Ppt_PasteKept(linked, linkState)
End Function

' Estado do vinculo do grafico colado: 1 vinculado, 0 colado sem vinculo, -1 desconhecido (leitura falhou). Nunca propaga erro.
Public Function Ppt_LinkState(shp As Object) As Long
    Dim v As Variant
    Ppt_LinkState = -1
    If shp Is Nothing Then Exit Function
    On Error Resume Next
    v = shp.Chart.ChartData.IsLinked
    If Err.Number = 0 Then
        If CBool(v) Then Ppt_LinkState = 1 Else Ppt_LinkState = 0
    End If
    Err.Clear
End Function

' Aviso de sucesso. Grafico que o PowerPoint colou sem o vinculo (linkState = 0) diz isso em vez de "vinculado".
Public Function Ppt_DoneText(slideIndex As Long, presName As String, linked As Boolean, linkState As Long) As String
    Dim how As String
    If Not linked Then
        how = "como imagem."
    ElseIf linkState = 0 Then
        how = "como gráfico do PowerPoint, mas sem vínculo com a pasta (o PowerPoint não manteve o vínculo): ele não acompanha as alterações dos dados."
    Else
        how = "como gráfico vinculado."
    End If
    Ppt_DoneText = "Gráfico copiado para o slide " & slideIndex & " de '" & presName & "' " & how
End Function

Public Function Ppt_LinkedPrecheck(wb As Workbook) As String
    If Len(wb.Path) = 0 Then Ppt_LinkedPrecheck = PPT_ERR_UNSAVED
End Function

' Falha esperada (PowerPoint ausente, area de transferencia ocupada): vira aviso proprio, sem a moldura de "erro do ThinQQ".
Public Function Ppt_IsKnownFailure(errNum As Long) As Boolean
    Ppt_IsKnownFailure = (errNum = 429 Or errNum = PPT_CLIPBRD_E_CANT_OPEN Or errNum = PPT_RPC_E_CALL_REJECTED Or errNum = PPT_RPC_E_RETRYLATER)
End Function

' PowerPoint ocupado (caixa de dialogo aberta, texto em edicao): a chamada pode ser repetida mais tarde.
Public Function Ppt_IsBusyError(errNum As Long) As Boolean
    Ppt_IsBusyError = (errNum = PPT_RPC_E_CALL_REJECTED Or errNum = PPT_RPC_E_RETRYLATER)
End Function

Public Function Ppt_ErrorText(errNum As Long, errDesc As String) As String
    If errNum = 429 Then
        Ppt_ErrorText = PPT_ERR_NOT_INSTALLED
    ElseIf errNum = PPT_CLIPBRD_E_CANT_OPEN Then
        Ppt_ErrorText = PPT_ERR_CLIPBOARD
    ElseIf errNum = PPT_RPC_E_CALL_REJECTED Or errNum = PPT_RPC_E_RETRYLATER Then
        Ppt_ErrorText = PPT_ERR_BUSY
    Else
        Ppt_ErrorText = "Não foi possível colar no PowerPoint: " & errDesc & " (" & errNum & ")"
    End If
End Function

Private Sub Ppt_PlaceShape(shp As Object, slide As Object)
    Dim sw As Double, sh As Double, w As Double, h As Double
    Ppt_Hook "place"
    sw = slide.Parent.PageSetup.SlideWidth
    sh = slide.Parent.PageSetup.SlideHeight
    SafeSlideSize sw, sh                             ' a centralizacao usa o mesmo slide que Ppt_FitSize
    Ppt_FitSize CDbl(shp.Width), CDbl(shp.Height), sw, sh, w, h
    shp.LockAspectRatio = -1
    shp.Width = w
    shp.Height = h
    shp.Left = (sw - shp.Width) / 2
    shp.Top = (sh - shp.Height) / 2
End Sub

' Cola o que esta na area de transferencia no slide e ajusta. Transacional: se o ajuste falhar, a forma colada e removida do slide.
Private Function Ppt_PasteAndPlace(slide As Object) As Object
    Dim sr As Object, shp As Object, en As Long, ed As String
    Set sr = slide.Shapes.Paste
    On Error GoTo Fail
    Set shp = sr.Item(1)
    Ppt_PlaceShape shp, slide
    Set Ppt_PasteAndPlace = shp
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    sr.Delete
    On Error GoTo 0
    Err.Raise en, "modPowerPoint", ed
End Function

Private Sub Ppt_Pause(seconds As Single)
    Dim t As Single
    t = Timer
    Do While Timer >= t And Timer - t < seconds
        DoEvents
    Loop
End Sub

' Copia para a area de transferencia; com ela aberta por outro programa (CLIPBRD_E_CANT_OPEN) tenta de novo algumas vezes (comum com
' gerenciadores de area de transferencia); qualquer outro erro, ou o esgotar das tentativas, sobe ao chamador.
Private Sub Ppt_CopyToClipboard(cht As Chart, asPicture As Boolean)
    Dim i As Long, en As Long, ed As String
    For i = 1 To PPT_CLIP_TRIES
        On Error Resume Next
        If asPicture Then cht.CopyPicture xlScreen, xlPicture Else cht.Parent.Copy
        en = Err.Number: ed = Err.Description
        On Error GoTo 0
        If en <> PPT_CLIPBRD_E_CANT_OPEN Then Exit For
        If i < PPT_CLIP_TRIES Then Ppt_Pause PPT_CLIP_WAIT
    Next
    If en <> 0 Then Err.Raise en, "modPowerPoint", ed
End Sub

Public Function Ppt_PasteImage(cht As Chart, slide As Object) As Object
    Ppt_CopyToClipboard cht, True
    Set Ppt_PasteImage = Ppt_PasteAndPlace(slide)
End Function

Public Function Ppt_PasteChart(co As ChartObject, slide As Object) As Object
    Ppt_CopyToClipboard co.Chart, False
    Set Ppt_PasteChart = Ppt_PasteAndPlace(slide)
End Function
