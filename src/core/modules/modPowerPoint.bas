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
Private Const PPT_LAYOUT_BLANK As Long = 12     ' ppLayoutBlank
Private Const PPT_FIT As Double = 0.8           ' fracao do slide que o grafico ocupa (largura ou altura, a que limitar)
Private Const PPT_CLIPBRD_E_CANT_OPEN As Long = -2147221040     ' 0x800401D0: a area de transferencia esta aberta por outro programa
Private Const PPT_CLIP_TRIES As Long = 3        ' tentativas de copiar quando a area de transferencia esta ocupada
Private Const PPT_CLIP_WAIT As Single = 0.25    ' segundos entre as tentativas
Public Const PPT_ERR_NOT_INSTALLED As String = "O PowerPoint não está instalado neste computador (ou não pôde ser iniciado)."
Public Const PPT_ERR_UNSAVED As String = "Salve a pasta de trabalho antes de copiar o gráfico vinculado: o vínculo aponta para o arquivo salvo."
Public Const PPT_ERR_CLIPBOARD As String = "A área de transferência do Windows não está disponível (outro programa a mantém aberta). Tente novamente em instantes."
Public Const PPT_NOTE_DIRTY As String = "A pasta tem alterações não salvas: o PowerPoint lê o arquivo salvo nas atualizações do vínculo."

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

Public Function Ppt_IsRunning() As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = GetObject(, "PowerPoint.Application")
    Ppt_IsRunning = (Err.Number = 0) And (Not o Is Nothing)
    Err.Clear
End Function

Public Function Ppt_GetApp(create As Boolean) As Object
    Dim o As Object
    mCreatedApp = False
    If BUILD_CONFIG = "Test" Then Set Ppt_GetApp = mTestApp: Exit Function     ' testes: so a instancia injetada (sem ela, Nothing): nunca o PowerPoint do usuario
    On Error Resume Next
    Set o = GetObject(, "PowerPoint.Application")
    Err.Clear
    If o Is Nothing And create Then
        Set o = CreateObject("PowerPoint.Application")
        mCreatedApp = Not o Is Nothing
    End If
    Err.Clear
    On Error GoTo 0
    Set Ppt_GetApp = o
End Function

' True se a ultima Ppt_GetApp abriu um PowerPoint novo (nunca no build Test: a instancia injetada nao e nossa para encerrar).
Public Function Ppt_AppWasCreated() As Boolean
    Ppt_AppWasCreated = mCreatedApp
End Function

' Slide de destino (ver cabecalho). Propaga erros do PowerPoint.
Public Function Ppt_TargetSlide(app As Object) As Object
    Dim pres As Object, slide As Object
    On Error Resume Next
    If app.Windows.Count > 0 Then Set slide = app.ActiveWindow.View.Slide
    Err.Clear
    On Error GoTo 0
    If Not slide Is Nothing Then Set Ppt_TargetSlide = slide: Exit Function
    If app.Presentations.Count > 0 Then
        Set pres = app.Presentations(1)
        If pres.Slides.Count = 0 Then Set slide = pres.Slides.Add(1, PPT_LAYOUT_BLANK) Else Set slide = pres.Slides(pres.Slides.Count)
    Else
        ' msoTrue: com janela, para o usuario ver (no build Test sem janela: nao aparece na area de trabalho)
        Set pres = app.Presentations.Add(IIf(BUILD_CONFIG = "Test", 0, -1))
        Ppt_Hook "slide"
        Set slide = pres.Slides.Add(1, PPT_LAYOUT_BLANK)
    End If
    Set Ppt_TargetSlide = slide
End Function

' Tamanho que cabe em PPT_FIT do slide mantendo a proporcao (pura). Tamanho invalido fica como esta.
Public Sub Ppt_FitSize(w As Double, h As Double, slideW As Double, slideH As Double, ByRef outW As Double, ByRef outH As Double)
    Dim f As Double
    outW = w: outH = h
    If w <= 0 Or h <= 0 Or slideW <= 0 Or slideH <= 0 Then Exit Sub
    f = PPT_FIT * slideW / w
    If h * f > PPT_FIT * slideH Then f = PPT_FIT * slideH / h
    outW = w * f: outH = h * f
End Sub

Public Function Ppt_LinkedPrecheck(wb As Workbook) As String
    If Len(wb.Path) = 0 Then Ppt_LinkedPrecheck = PPT_ERR_UNSAVED
End Function

' Falha esperada (PowerPoint ausente, area de transferencia ocupada): vira aviso proprio, sem a moldura de "erro do ThinQQ".
Public Function Ppt_IsKnownFailure(errNum As Long) As Boolean
    Ppt_IsKnownFailure = (errNum = 429 Or errNum = PPT_CLIPBRD_E_CANT_OPEN)
End Function

Public Function Ppt_ErrorText(errNum As Long, errDesc As String) As String
    If errNum = 429 Then
        Ppt_ErrorText = PPT_ERR_NOT_INSTALLED
    ElseIf errNum = PPT_CLIPBRD_E_CANT_OPEN Then
        Ppt_ErrorText = PPT_ERR_CLIPBOARD
    Else
        Ppt_ErrorText = "Não foi possível colar no PowerPoint: " & errDesc & " (" & errNum & ")"
    End If
End Function

Private Sub Ppt_PlaceShape(shp As Object, slide As Object)
    Dim sw As Double, sh As Double, w As Double, h As Double
    Ppt_Hook "place"
    sw = slide.Parent.PageSetup.SlideWidth
    sh = slide.Parent.PageSetup.SlideHeight
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
