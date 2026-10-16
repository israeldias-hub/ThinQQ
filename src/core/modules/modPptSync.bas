Attribute VB_Name = "modPptSync"
Option Explicit
' Graficos ThinQQ colados no PowerPoint como grafico vinculado (Parte A, spec 2026-10-14 secao 4.2): etiquetas na forma do slide (TQQ_*) e
' sincronizacao. Dados pelo vinculo nativo (LinkFormat.Update: le a pasta ABERTA, mesmo com alteracoes nao salvas; fechada, abre um Excel
' -Embedding, ~4 s); as formas das anotacoes (tqq_ann_*) sao recalculadas pelo ThinQQ com o grafico do Excel no TAMANHO DO SLIDE e a foto
' delas e aplicada dentro do grafico do slide, NO LUGAR (animacao, ordem Z, nome e texto alternativo da forma do slide ficam). Medido (spec
' 3.1): com o mesmo tamanho a area interna dos dois aplicativos e identica; redimensionar o grafico do slide ESTICA as formas internas;
' LinkFormat.Update com a pasta de origem inexistente TRAVA o PowerPoint: quem decide e PptSync_SyncShape (so chama o Update se
' PptSync_SourceReachable) e PptSync_Presentation confere o arquivo antes de abrir a pasta e nunca grava a impressao digital de um grafico
' cujo Update foi ignorado (limites: SourceFullName ilegivel cai na etiqueta TQQ_BOOK; FileExists em compartilhamento de rede morto pode
' demorar; pasta no OneDrive/SharePoint (https) so e sincronizada se estiver aberta no Excel).
' Ligacao tardia com o PowerPoint (As Object, constantes numericas); sem Declare.
Public Const PPTSYNC_TAG_CHART As String = "TQQ_CHART"      ' id do grafico (ChartObject "tqq_" & id)
Public Const PPTSYNC_TAG_BOOK As String = "TQQ_BOOK"        ' caminho completo da pasta na colagem
Public Const PPTSYNC_TAG_SHEET As String = "TQQ_SHEET"      ' planilha do grafico na colagem (informativo)
Public Const PPTSYNC_TAG_VER As String = "TQQ_VER"          ' versao do formato das etiquetas
Public Const PPTSYNC_TAG_FP As String = "TQQ_FP"            ' impressao digital da ultima sincronizacao (PptSync_Fingerprint)
Public Const PPTSYNC_TAG_TITLE As String = "TQQ_TITLE"      ' "1" = o titulo do grafico do slide foi posto pelo ThinQQ (titulo ligado, 1.9.0)
Public Const PPTSYNC_TAG_AUTO As String = "TQQ_AUTO"        ' "1" = atualizar automaticamente (spec 2026-10-15 secao 5.1)
Public Const PPTSYNC_TAG_INDEX As String = "TQQ_AUTOIDX"   ' etiqueta da APRESENTACAO: indice das formas automaticas (PptIdx_*)
Public Const PPTIDX_MAX_ENTRIES As Long = 2000             ' PptIdx_Parse: entradas alem disto sao ignoradas
Public Const PPTIDX_MAX_LEN As Long = 1000000              ' PptIdx_Parse: etiqueta maior que isto (caracteres) e recusada inteira
Private Const PPTIDX_FS As String = "|"                    ' indice: separador de campos (PptIdx_*)
Private Const PPTIDX_ES As String = vbLf                   ' indice: separador de entradas
Public Const PPTSYNC_TAG_KIND As String = "TQQ_KIND"        ' tipo do grafico (a faixa do PowerPoint le sem chamar o Excel; spec 14.5)
Public Const PPTSYNC_TAG_OPTS As String = "TQQ_OPTS"        ' estado das opcoes (PptSync_OptsText) para getPressed/getEnabled do .ppam
Public Const CORE_PPT_API As Long = 5                      ' 1 = Core_PptSyncPresentation; 2 = Core_PptSyncShapes, Core_PptGetFlag/SetFlag (1.6.0);
                                                           ' 3 = Core_PptInsertChart, Core_PptEditData, Core_PptDataSave (1.8.0);
                                                           ' 4 = Core_PptCommand, Core_PptCategories (1.9.0); 5 = Core_PptLinkText (1.9.0)
Public Const PPTSYNC_NONE As String = "Esta apresentação não tem gráficos do ThinQQ vinculados (copiados com 'Copiar como gráfico vinculado')."
Public Const PPTSYNC_NOT_FOUND As String = "Os gráficos selecionados não foram encontrados na apresentação (renomeados?): selecione-os de novo."
Public Const PPTSYNC_NO_PPT As String = "Abra a apresentação no PowerPoint e use de novo 'Atualizar gráficos no PowerPoint'."
Public Const PPTSYNC_UNKNOWN_SRC As String = "(origem desconhecida)"      ' forma marcada sem SourceFullName e sem TQQ_BOOK
Public Const PPTSYNC_C_PROCEED As Long = 0                   ' PptSync_Classify: segue para a pasta de origem
Public Const PPTSYNC_C_UNLINKED As Long = 1                  ' colado sem vinculo ("Sem vinculo com o Excel")
Public Const PPTSYNC_C_UNKNOWN As Long = 2                   ' estado do vinculo ilegivel (falha)
Public Const PPTSYNC_C_NOPATH As Long = 3                    ' vinculado, mas sem caminho de origem (origem desconhecida)
Public Const PPTSYNC_R_OK As Long = 1                        ' resultado de PptSync_SyncOne e PptText_Sync
Public Const PPTSYNC_R_SAME As Long = 2
Public Const PPTSYNC_R_NOCHART As Long = 3
Public Const PPTSYNC_R_FAIL As Long = 4
Public Const PPTSYNC_R_SKIPPED As Long = 5
Public Const PPTSYNC_R_EDITED As Long = 6                    ' texto vinculado editado fora do campo (nada tocado)
Private Const MSO_AUTOSHAPE As Long = 1
Private Const MSO_LINE As Long = 9
Private Const MSO_TEXTBOX As Long = 17
Private Const SN_NAME As Long = 0                           ' campos da foto de uma forma (PptSync_Snapshot)
Private Const SN_LEFT As Long = 1
Private Const SN_TOP As Long = 2
Private Const SN_WIDTH As Long = 3
Private Const SN_HEIGHT As Long = 4
Private Const SN_VISIBLE As Long = 5
Private Const SN_HFLIP As Long = 6
Private Const SN_VFLIP As Long = 7
Private Const SN_TEXT As Long = 8
Private Const SN_HASTEXT As Long = 9
Private Const PPTSYNC_LAYOUT_TOL As Double = 2              ' pt: area interna do slide diferente da do Excel alem disso = log
Private mTestFail As String                                 ' gancho de teste: ponto de falha injetada (so no build Test)
Private mTestSkip As Long                                 ' gancho de teste: quantos acertos do ponto saltar antes de falhar (PptSync_SetFailAfterForTest)
Private mUpdateCount As Long                                ' chamadas de LinkFormat.Update feitas por PptSync_SyncShape (lido so no build Test)
Private mTestFailNum As Long                                ' gancho de teste: numero do erro injetado (0 = vbObjectError + 9992; PptSync_SetFailNumberForTest)
Private mTestKeepEventsOnClose As Boolean                   ' gancho de teste: PptSync_CloseOpened deixa os eventos como estao (prova que so as macros desligadas bastam)

' Gancho de teste: o proximo ponto nomeado ("update", "snapshot", "apply") levanta erro (uso unico). Inerte fora do build Test.
' Gancho de teste: o proximo PptSync_CloseOpened NAO desliga os eventos (zerado por PptSync_ResetTestHooks). Inerte fora do build Test.
Public Sub PptSync_KeepEventsOnCloseForTest()
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestKeepEventsOnClose = True
End Sub

Public Sub PptSync_SetFailForTest(point As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestFail = point
    mTestSkip = 0
    mTestFailNum = 0
End Sub

' Gancho de teste: como PptSync_SetFailForTest, mas salta os primeiros skipFirst acertos do ponto (a falha cai na (skipFirst+1)-esima vez).
Public Sub PptSync_SetFailAfterForTest(point As String, skipFirst As Long)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestFail = point
    mTestSkip = skipFirst
    mTestFailNum = 0
End Sub

' Gancho de teste: como PptSync_SetFailForTest, mas a falha injetada tem o numero errNum (ex.: PowerPoint ocupado, -2147418111). Uso unico.
Public Sub PptSync_SetFailNumberForTest(point As String, errNum As Long)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestFail = point
    mTestSkip = 0
    mTestFailNum = errNum
End Sub

Public Sub PptSync_ResetTestHooks()
    mTestKeepEventsOnClose = False
    mTestFail = ""
    mTestFailNum = 0
    mTestSkip = 0
    mUpdateCount = 0
End Sub

' Gancho de teste: quantas vezes PptSync_SyncShape chamou LinkFormat.Update desde o ultimo PptSync_ResetTestHooks (0 fora do build Test).
Public Function PptSync_UpdateCountForTest() As Long
    If BUILD_CONFIG <> "Test" Then Exit Function
    PptSync_UpdateCountForTest = mUpdateCount
End Function

Private Sub PptSync_Hook(point As String)
    Dim n As Long
    If BUILD_CONFIG <> "Test" Then Exit Sub
    If mTestFail <> point Then Exit Sub
    If mTestSkip > 0 Then mTestSkip = mTestSkip - 1: Exit Sub
    mTestFail = ""
    n = mTestFailNum: mTestFailNum = 0
    If n = 0 Then n = vbObjectError + 9992
    Err.Raise n, "modPptSync", "falha injetada (" & point & ")"
End Sub

' Ponto de falha injetada para outros modulos (modPptAuto: leituras de etiqueta). "" = nenhum. Inerte fora do build Test.
Public Sub PptSync_TestPoint(point As String)
    If Len(point) = 0 Then Exit Sub
    PptSync_Hook point
End Sub

' Gancho de teste "unreachable" (uso unico): faz PptSync_SyncShape ver a origem como inalcancavel (a pasta sumiu entre a conferencia de
' PptSync_Presentation e o Update). Nao levanta erro. Sempre False fora do build Test.
Private Function PptSync_TakeUnreachableHook() As Boolean
    If BUILD_CONFIG <> "Test" Then Exit Function
    If mTestFail <> "unreachable" Then Exit Function
    mTestFail = ""
    PptSync_TakeUnreachableHook = True
End Function

' Etiquetas da forma colada (grafico vinculado): o que o suplemento do PowerPoint e PptSync_Presentation leem para achar o grafico de origem.
Public Sub PptSync_TagShape(shp As Object, co As ChartObject)
    shp.Tags.Add PPTSYNC_TAG_CHART, Registry_IdFromChart(co.Chart)
    shp.Tags.Add PPTSYNC_TAG_BOOK, Registry_Workbook(co.Chart).FullName
    shp.Tags.Add PPTSYNC_TAG_SHEET, co.Parent.Name
    shp.Tags.Add PPTSYNC_TAG_VER, "1"
End Sub

' Valor de uma etiqueta ("" se nao existe ou a leitura falha). Nunca propaga erro.
Public Function PptSync_Tag(shp As Object, key As String) As String
    On Error Resume Next
    PptSync_Tag = CStr(shp.Tags.Item(key))
    Err.Clear
End Function

' Caminho da pasta de origem: LinkFormat.SourceFullName ate o primeiro "!" (o grafico vinculado devolve o caminho puro; objeto OLE devolve
' "caminho!planilha!..."); sem vinculo legivel, a etiqueta TQQ_BOOK. Pura.
Public Function PptSync_SourcePath(linkFull As String, tagBook As String) As String
    Dim s As String, p As Long
    s = Trim$(linkFull)
    p = InStr(s, "!")
    If p > 0 Then s = Left$(s, p - 1)
    If Len(s) = 0 Then s = Trim$(tagBook)
    PptSync_SourcePath = s
End Function

' A pasta de origem pode ser lida pelo vinculo sem travar o PowerPoint (Update com a origem inexistente trava, spec 3.1): aberta neste Excel
' (FullName igual) ou, se for caminho local/de rede, existente no disco. Endereco http(s) (OneDrive/SharePoint) so conta aberta. Nunca
' propaga erro.
Public Function PptSync_SourceReachable(path As String) As Boolean
    Dim wb As Workbook, names As New Collection
    On Error GoTo Done
    For Each wb In Application.Workbooks
        names.Add wb.FullName
    Next
    PptSync_SourceReachable = PptSync_ReachableAmong(path, names)
Done:
End Function

' Regra de PptSync_SourceReachable com a lista dos FullName das pastas abertas (names): path igual a um deles (sem diferenciar maiusculas;
' vale tambem para o endereco https do OneDrive/SharePoint) ou, se nao for http(s), arquivo existente no disco. Nunca propaga erro.
Public Function PptSync_ReachableAmong(path As String, names As Collection) As Boolean
    Dim n As Variant, p As String
    On Error GoTo Done
    p = Trim$(path)
    If Len(p) = 0 Then Exit Function
    For Each n In names
        If StrComp(CStr(n), p, vbTextCompare) = 0 Then PptSync_ReachableAmong = True: Exit Function
    Next
    If LCase$(Left$(p, 7)) = "http://" Or LCase$(Left$(p, 8)) = "https://" Then Exit Function
    PptSync_ReachableAmong = CBool(CreateObject("Scripting.FileSystemObject").FileExists(p))
Done:
End Function

' LinkFormat.SourceFullName da forma ("" se nao ha vinculo legivel). Nunca propaga erro.
Private Function PptSync_LinkSource(shp As Object) As String
    On Error Resume Next
    PptSync_LinkSource = CStr(shp.LinkFormat.SourceFullName)
    Err.Clear
End Function

' Foto das formas das anotacoes (prefixo tqq_ann_) do grafico do Excel: nome, caixa, visivel, giros e texto. Chave da colecao = nome.
Public Function PptSync_Snapshot(cht As Chart) As Collection
    Dim out As New Collection, i As Long, s As Shape, txt As String, hasTxt As Boolean
    For i = 1 To cht.Shapes.Count
        Set s = cht.Shapes(i)
        If Left$(s.Name, Len(ANN_PREFIX)) = ANN_PREFIX Then
            txt = "": hasTxt = False
            If s.Type <> MSO_LINE Then
                On Error Resume Next
                txt = s.TextFrame2.TextRange.Text
                hasTxt = (Err.Number = 0)
                Err.Clear
                On Error GoTo 0
            End If
            out.Add Array(s.Name, CDbl(s.Left), CDbl(s.Top), CDbl(s.Width), CDbl(s.Height), CLng(s.Visible), CLng(s.HorizontalFlip), CLng(s.VerticalFlip), txt, hasTxt), s.Name
        End If
    Next
    Set PptSync_Snapshot = out
End Function

' Aplica a foto no grafico do slide (pc): apaga as tqq_ann_* do slide que nao existem mais no Excel, clona as que faltam (src = grafico do
' Excel, para o formato) e poe caixa, giros, texto e visibilidade em todas, deslocadas por (dx, dy). Devolve quantas formas clonou.
Public Function PptSync_Apply(pc As Object, snaps As Collection, src As Chart, dx As Double, dy As Double) As Long
    Dim i As Long, nm As String, t As Object, sn As Variant, keep As String
    keep = "|"
    For Each sn In snaps
        keep = keep & sn(SN_NAME) & "|"
    Next
    For i = pc.Shapes.Count To 1 Step -1
        nm = CStr(pc.Shapes(i).Name)
        If Left$(nm, Len(ANN_PREFIX)) = ANN_PREFIX And InStr(keep, "|" & nm & "|") = 0 Then pc.Shapes(i).Delete
    Next
    For Each sn In snaps
        Set t = Nothing
        On Error Resume Next
        Set t = pc.Shapes(CStr(sn(SN_NAME)))
        Err.Clear
        On Error GoTo 0
        If t Is Nothing Then
            Set t = PptSync_Clone(pc, src.Shapes(CStr(sn(SN_NAME))))
            If Not t Is Nothing Then PptSync_Apply = PptSync_Apply + 1
        End If
        If Not t Is Nothing Then PptSync_Place t, sn, dx, dy
    Next
End Function

Private Sub PptSync_Place(t As Object, sn As Variant, dx As Double, dy As Double)
    If CBool(sn(SN_HASTEXT)) Then
        If CStr(t.TextFrame2.TextRange.Text) <> CStr(sn(SN_TEXT)) Then t.TextFrame2.TextRange.Text = CStr(sn(SN_TEXT))
    End If
    If (CLng(t.HorizontalFlip) = -1) <> (CLng(sn(SN_HFLIP)) = -1) Then t.Flip 0          ' msoFlipHorizontal
    If (CLng(t.VerticalFlip) = -1) <> (CLng(sn(SN_VFLIP)) = -1) Then t.Flip 1            ' msoFlipVertical
    t.Left = CSng(sn(SN_LEFT) + dx)
    t.Top = CSng(sn(SN_TOP) + dy)
    t.Width = CSng(sn(SN_WIDTH))
    t.Height = CSng(sn(SN_HEIGHT))
    If CLng(t.Visible) <> CLng(sn(SN_VISIBLE)) Then t.Visible = CLng(sn(SN_VISIBLE))
End Sub

' Cria no grafico do slide a forma que so existe no Excel (seta inserida depois da colagem): linha, caixa de texto ou autoforma, com o
' formato copiado (o melhor esforco: um formato que nao copia so e registrado). Outros tipos: Nothing (registrado).
Private Function PptSync_Clone(pc As Object, s As Shape) As Object
    Dim t As Object
    Select Case s.Type
        Case MSO_LINE
            Set t = pc.Shapes.AddLine(s.Left, s.Top, s.Left + s.Width, s.Top + s.Height)
            PptSync_CopyLine s, t
        Case MSO_TEXTBOX
            Set t = pc.Shapes.AddTextbox(1, s.Left, s.Top, s.Width, s.Height)        ' msoTextOrientationHorizontal
            PptSync_CopyText s, t
            PptSync_CopyFillAndBorder s, t
        Case MSO_AUTOSHAPE
            Set t = pc.Shapes.AddShape(s.AutoShapeType, s.Left, s.Top, s.Width, s.Height)
            PptSync_CopyFillAndBorder s, t
            PptSync_CopyText s, t
        Case Else
            Errors_Log "PptSync_Clone: tipo de forma nao copiado (" & s.Type & ")" & vbTab & s.Name
            Exit Function
    End Select
    t.Name = s.Name
    t.AlternativeText = s.AlternativeText
    Set PptSync_Clone = t
End Function

Private Sub PptSync_CopyLine(s As Shape, t As Object)
    On Error Resume Next
    With t.Line
        .Visible = s.Line.Visible
        .ForeColor.RGB = s.Line.ForeColor.RGB
        .Weight = s.Line.Weight
        .DashStyle = s.Line.DashStyle
        .BeginArrowheadStyle = s.Line.BeginArrowheadStyle
        .EndArrowheadStyle = s.Line.EndArrowheadStyle
        .BeginArrowheadLength = s.Line.BeginArrowheadLength
        .EndArrowheadLength = s.Line.EndArrowheadLength
        .BeginArrowheadWidth = s.Line.BeginArrowheadWidth
        .EndArrowheadWidth = s.Line.EndArrowheadWidth
    End With
    If Err.Number <> 0 Then Errors_Log "PptSync_CopyLine " & s.Name & vbTab & Err.Number & vbTab & Err.Description
End Sub

Private Sub PptSync_CopyText(s As Shape, t As Object)
    On Error Resume Next
    With t.TextFrame2
        .WordWrap = s.TextFrame2.WordWrap
        .AutoSize = s.TextFrame2.AutoSize
        .MarginLeft = s.TextFrame2.MarginLeft: .MarginRight = s.TextFrame2.MarginRight
        .MarginTop = s.TextFrame2.MarginTop: .MarginBottom = s.TextFrame2.MarginBottom
        .TextRange.Text = s.TextFrame2.TextRange.Text
        .TextRange.Font.Name = s.TextFrame2.TextRange.Font.Name
        .TextRange.Font.Size = s.TextFrame2.TextRange.Font.Size
        .TextRange.Font.Bold = s.TextFrame2.TextRange.Font.Bold
        .TextRange.Font.Fill.ForeColor.RGB = s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
        .TextRange.ParagraphFormat.Alignment = s.TextFrame2.TextRange.ParagraphFormat.Alignment
    End With
    If Err.Number <> 0 Then Errors_Log "PptSync_CopyText " & s.Name & vbTab & Err.Number & vbTab & Err.Description
End Sub

Private Sub PptSync_CopyFillAndBorder(s As Shape, t As Object)
    On Error Resume Next
    ' medido: numa caixa de texto do grafico do slide, Fill.Visible = msoTrue falha ("Erro nao especificado"); a cor liga o preenchimento
    If s.Fill.Visible = msoTrue Then t.Fill.ForeColor.RGB = s.Fill.ForeColor.RGB Else t.Fill.Visible = msoFalse
    t.Line.Visible = s.Line.Visible
    If s.Line.Visible = msoTrue Then
        t.Line.ForeColor.RGB = s.Line.ForeColor.RGB
        t.Line.Weight = s.Line.Weight
        t.Line.DashStyle = s.Line.DashStyle
    End If
    If Err.Number <> 0 Then Errors_Log "PptSync_CopyFillAndBorder " & s.Name & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Forca o Excel a refazer a geometria do grafico (pontos, rotulos, area interna): exportar e o unico jeito no Excel invisivel (o ForceLayout
' dos testes faz o mesmo). PNG com nome unico na pasta temporaria do Windows, apagado mesmo se a exportacao falhar; sem Dir$ (zeraria um
' laco Dir$ de quem chama). Nunca propaga erro.
Public Sub PptSync_ForceLayout(cht As Chart)
    Dim fso As Object, p As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then cht.Refresh: Err.Clear: Exit Sub
    p = fso.BuildPath(fso.GetSpecialFolder(2), "thinqq_pptsync_" & fso.GetBaseName(fso.GetTempName()) & ".png")     ' 2 = TemporaryFolder
    cht.Export p
    Err.Clear
    If fso.FileExists(p) Then fso.DeleteFile p, True
    Err.Clear
End Sub

' Sincroniza UM grafico do slide (shp: forma do PowerPoint com o grafico vinculado) com o grafico ThinQQ co: (1) dados pelo vinculo nativo
' (so com o vinculo ativo); (2) o ChartObject vai para o tamanho da forma do slide, o ThinQQ recalcula as anotacoes (Core_SafeRefresh) e a foto
' delas e aplicada no grafico do slide com o deslocamento das areas internas (medido 0); (3) o ChartObject volta ao tamanho original e a pasta
' volta a "salva" se estava. Levanta erro (quem chama registra e conta); o tamanho e restaurado em qualquer caminho. updateData = False pula o
' passo (1) (colagem: os dados acabaram de ser colados). O Update so roda com a pasta de origem alcancavel (PptSync_SourceReachable); quando
' ele e ignorado por isso, updateSkipped volta True (False em todos os outros casos) e quem chama NAO pode gravar a TQQ_FP nem contar o
' grafico como atualizado (os dados do slide ficaram como estavam).
Public Sub PptSync_SyncShape(shp As Object, co As ChartObject, Optional updateData As Boolean = True, Optional ByRef updateSkipped As Boolean)
    Dim cht As Chart, wb As Workbook, w0 As Double, h0 As Double, saved0 As Boolean, resized As Boolean, prevSU As Boolean, prevEv As Boolean
    Dim snaps As Collection, pc As Object, dx As Double, dy As Double, en As Long, ed As String, src As String
    updateSkipped = False
    Set cht = co.Chart: Set wb = Registry_Workbook(cht)
    saved0 = wb.Saved: w0 = co.Width: h0 = co.Height
    prevSU = Application.ScreenUpdating: prevEv = Application.EnableEvents
    On Error GoTo Fail
    If updateData Then
        If Ppt_LinkState(shp) = 1 Then
            src = PptSync_SourcePath(PptSync_LinkSource(shp), PptSync_Tag(shp, PPTSYNC_TAG_BOOK))
            If PptSync_SourceReachable(src) And Not PptSync_TakeUnreachableHook() Then
                shp.LinkFormat.Update
                mUpdateCount = mUpdateCount + 1
            Else
                updateSkipped = True
                Errors_Log "PptSync: pasta de origem nao encontrada; dados do slide nao atualizados" & vbTab & src & vbTab & co.Name
            End If
        End If
    End If
    PptSync_Hook "update"
    Application.ScreenUpdating = False
    Application.EnableEvents = False                    ' a janela do redimensionamento nao dispara os vigias do grafico (cada Resize refazia o Core_SafeRefresh)
    resized = True
    co.Width = CDbl(shp.Width): co.Height = CDbl(shp.Height)
    PptSync_ForceLayout cht
    Core_SafeRefresh cht
    PptSync_ForceLayout cht
    PptSync_Hook "snapshot"
    Set snaps = PptSync_Snapshot(cht)
    PptSync_Hook "apply"
    Set pc = shp.Chart
    PptSync_ApplyTitle shp, pc, cht                     ' antes de medir as areas internas: o titulo ocupa espaco
    dx = CDbl(pc.PlotArea.InsideLeft) - cht.PlotArea.InsideLeft
    dy = CDbl(pc.PlotArea.InsideTop) - cht.PlotArea.InsideTop
    If Abs(CDbl(pc.PlotArea.InsideWidth) - cht.PlotArea.InsideWidth) > PPTSYNC_LAYOUT_TOL Or Abs(CDbl(pc.PlotArea.InsideHeight) - cht.PlotArea.InsideHeight) > PPTSYNC_LAYOUT_TOL Then
        Errors_Log "PptSync: area interna do slide diferente da do Excel (grafico editado no PowerPoint?)" & vbTab & CStr(shp.Name)
    End If
    PptSync_Apply pc, snaps, cht, dx, dy
    PptSync_Restore co, w0, h0, wb, saved0
    Application.ScreenUpdating = prevSU: Application.EnableEvents = prevEv
    Exit Sub
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    If resized Then PptSync_Restore co, w0, h0, wb, saved0
    Application.ScreenUpdating = prevSU: Application.EnableEvents = prevEv
    On Error GoTo 0
    Err.Raise en, "PptSync_SyncShape", ed
End Sub

' Titulo ligado a uma celula (1.9.0; revisao final E, I2): MEDIDO, o LinkFormat.Update do PowerPoint so traz os dados, nunca o titulo do
' grafico do Excel. Com o titulo ligado pelo ThinQQ (opcao titleref), o texto e copiado para o grafico do slide e a forma ganha a marca
' TQQ_TITLE; ligacao removida no Excel: o titulo posto pelo ThinQQ sai do slide. Um titulo do slide que o ThinQQ nao pos nunca e tocado.
' Levanta erro (PowerPoint ocupado etc.; quem chama trata como a falha da forma).
Private Sub PptSync_ApplyTitle(shp As Object, pc As Object, cht As Chart)
    Dim t As String
    If Len(Registry_GetOption(cht, "titleref", "")) > 0 And CBool(cht.HasTitle) Then
        t = CStr(cht.ChartTitle.Text)
        If Not CBool(pc.HasTitle) Then pc.HasTitle = True
        If CStr(pc.ChartTitle.Text) <> t Then pc.ChartTitle.Text = t
        If PptSync_Tag(shp, PPTSYNC_TAG_TITLE) <> "1" Then shp.Tags.Add PPTSYNC_TAG_TITLE, "1"
    ElseIf PptSync_Tag(shp, PPTSYNC_TAG_TITLE) = "1" Then
        If CBool(pc.HasTitle) Then pc.HasTitle = False
        PptSync_TagDelete shp, PPTSYNC_TAG_TITLE
    End If
End Sub

Private Sub PptSync_Restore(co As ChartObject, w0 As Double, h0 As Double, wb As Workbook, saved0 As Boolean)
    If Abs(co.Width - w0) > 0.01 Or Abs(co.Height - h0) > 0.01 Then
        co.Width = w0: co.Height = h0
        PptSync_ForceLayout co.Chart                        ' como na ida: o Excel invisivel so refaz a geometria ao desenhar
        Core_SafeRefresh co.Chart
        PptSync_ForceLayout co.Chart
    End If
    If saved0 Then wb.Saved = True
End Sub

' Impressao digital do que o slide mostraria: SHA-256 do bloco inteiro do grafico (valores e a linha de metadados com as opcoes), dos nomes e
' textos alternativos das anotacoes e do tamanho da forma do slide. Igual = nada a atualizar (PptSync_Presentation sem "forcar").
Public Function PptSync_Fingerprint(co As ChartObject, shp As Object) As String
    Dim blk As Range, v As Variant, s As String, r As Long, c As Long, i As Long, b() As Byte, nm As String
    Set blk = Registry_Block(co.Chart)
    If Not blk Is Nothing Then
        v = blk.Value
        If IsArray(v) Then
            For r = LBound(v, 1) To UBound(v, 1)
                For c = LBound(v, 2) To UBound(v, 2)
                    ' CStr segue a localidade do Windows: trocar a localidade (ou abrir em outra maquina) muda a impressao e so forca uma atualizacao
                    If IsError(v(r, c)) Then s = s & "#" Else s = s & CStr(v(r, c))
                    s = s & Chr$(31)
                Next
            Next
        ElseIf Not IsError(v) Then
            s = CStr(v)
        End If
    End If
    For i = 1 To co.Chart.Shapes.Count
        nm = co.Chart.Shapes(i).Name
        If Left$(nm, Len(ANN_PREFIX)) = ANN_PREFIX Then s = s & nm & "=" & co.Chart.Shapes(i).AlternativeText & Chr$(30)
    Next
    s = s & "|" & Format$(CDbl(shp.Width), "0.0") & "x" & Format$(CDbl(shp.Height), "0.0")
    s = s & "|t=" & PptSync_TitleText(co.Chart)              ' titulo (ligado a uma celula, 1.9.0): mudou = atualizar (revisao final E, I2)
    b = StrConv(s, vbFromUnicode)
    PptSync_Fingerprint = Sha256_Bytes(b, UBound(b) - LBound(b) + 1)
End Function

' Texto do titulo do grafico ("" sem titulo ou ilegivel). Nunca propaga erro.
Private Function PptSync_TitleText(c As Chart) As String
    On Error Resume Next
    If CBool(c.HasTitle) Then PptSync_TitleText = CStr(c.ChartTitle.Text)
    Err.Clear
End Function

' ---- Indice das formas automaticas (1.6.0). Formato (igual no .ppam: PptTqq_Idx*): entradas separadas por quebra de linha (vbLf), campos
' por "|": SlideID, nome da forma, pasta de origem. Medido no PowerPoint real (revisao final da C): ao salvar o .pptx a etiqueta perde
' Chr(30)/Chr(31) (caracteres proibidos no XML) e troca Tab por espaco e CR por LF; vbLf e "|" voltam iguais depois de reabrir. Puras e
' totais: a etiqueta vem da apresentacao (de quem a fez), entao lixo nunca levanta erro, o SlideID so vale com 1 a 9 algarismos, entradas alem
' de PPTIDX_MAX_ENTRIES sao ignoradas e uma etiqueta maior que PPTIDX_MAX_LEN e recusada inteira. Os separadores nunca entram num campo
' (PptIdx_SafeField). ----

' Campo do indice sem os separadores nem os caracteres que o .pptx salvo muda ("|", LF, CR, Tab, Chr(30), Chr(31) viram "_"). Pura.
Public Function PptIdx_SafeField(s As String) As String
    Dim t As String
    t = Replace(Replace(Replace(s, PPTIDX_FS, "_"), vbLf, "_"), vbCr, "_")
    PptIdx_SafeField = Replace(Replace(Replace(t, vbTab, "_"), Chr$(30), "_"), Chr$(31), "_")
End Function

Public Function PptIdx_Entry(slideId As Long, shapeName As String, book As String) As String
    PptIdx_Entry = CStr(slideId) & PPTIDX_FS & PptIdx_SafeField(shapeName) & PPTIDX_FS & PptIdx_SafeField(book)
End Function

Public Function PptIdx_Parse(idx As String) As Collection
    Dim out As New Collection, parts() As String, f() As String, i As Long
    Set PptIdx_Parse = out
    On Error GoTo Done
    If Len(idx) = 0 Or Len(idx) > PPTIDX_MAX_LEN Then Exit Function
    parts = Split(idx, PPTIDX_ES)
    For i = LBound(parts) To UBound(parts)
        If out.Count >= PPTIDX_MAX_ENTRIES Then Exit For
        If Len(parts(i)) > 0 Then
            f = Split(parts(i), PPTIDX_FS)
            If UBound(f) = 2 Then
                If Len(f(0)) > 0 And Len(f(0)) < 10 And Len(f(1)) > 0 Then
                    If Not f(0) Like "*[!0-9]*" Then out.Add Array(CLng(f(0)), f(1), f(2))
                End If
            End If
        End If
    Next
Done:
End Function

Public Function PptIdx_Join(entries As Collection) As String
    Dim e As Variant, a() As String, n As Long
    If entries.Count = 0 Then Exit Function
    ReDim a(0 To entries.Count - 1)
    For Each e In entries
        a(n) = PptIdx_Entry(CLng(e(0)), CStr(e(1)), CStr(e(2)))
        n = n + 1
    Next
    PptIdx_Join = Join(a, PPTIDX_ES)
End Function

' Tira a forma (slideId, nome) do indice e, com turnOn, a acrescenta no fim. Pura.
Public Function PptIdx_Set(idx As String, slideId As Long, shapeName As String, book As String, turnOn As Boolean) As String
    Dim out As New Collection, e As Variant, nm As String
    nm = PptIdx_SafeField(shapeName)
    For Each e In PptIdx_Parse(idx)
        If Not (CLng(e(0)) = slideId And StrComp(CStr(e(1)), nm, vbBinaryCompare) = 0) Then out.Add e
    Next
    If turnOn Then out.Add Array(slideId, nm, book)
    PptIdx_Set = PptIdx_Join(out)
End Function

' Apaga uma etiqueta (forma ou apresentacao); inexistente nao e erro. Nunca propaga erro.
Public Sub PptSync_TagDelete(o As Object, key As String)
    On Error Resume Next
    o.Tags.Delete key
    Err.Clear
End Sub

' Tipo do elemento do ThinQQ na forma do slide: "chart" (TQQ_CHART), "text" (TQQ_TEXT = "1": texto vinculado a celula, 1.9.0); "" = nao e
' do ThinQQ. A versao seguinte acrescenta "table". Nunca propaga erro.
Public Function PptSync_ItemKind(shp As Object) As String
    If Len(PptSync_Tag(shp, PPTSYNC_TAG_CHART)) > 0 Then
        PptSync_ItemKind = "chart"
    ElseIf PptSync_Tag(shp, PPTTEXT_TAG) = "1" Then
        PptSync_ItemKind = "text"
    End If
End Function

' Pasta de origem do elemento (vinculo nativo ou TQQ_BOOK). Nunca propaga erro.
Public Function PptSync_ItemBook(shp As Object) As String
    PptSync_ItemBook = PptSync_SourcePath(PptSync_LinkSource(shp), PptSync_Tag(shp, PPTSYNC_TAG_BOOK))
End Function

' Estado das opcoes do grafico para a faixa do PowerPoint ler SEM chamar o Excel (spec 14.5): "labels=1;totals=0;connectors=1;numfmt=int;
' legend=right;caps=labels,totals,...,". Pura sobre o grafico. Os valores vem do registro do grafico (que a pasta pode trazer alterado): so
' letras e digitos passam (PptSync_OptValue), entao o texto nunca tem ";", "=", ",", quebras de linha nem Chr(30)/Chr(31) fora da estrutura.
Public Function PptSync_OptsText(c As Chart) As String
    Dim k As Variant, s As String, caps As String, kind As String, nSeries As Long
    kind = Registry_ChartType(c)                         ' lido uma vez (revisao final G-II, M9): a mesma regra de Kinds_SupportsChart
    If Kinds_Family(kind) = "waterfall" Then nSeries = Waterfall_SeriesCount(c)
    For Each k In Array("labels", "totals", "connectors")
        s = s & k & "=" & IIf(Registry_GetOption(c, CStr(k), "0") = "1", "1", "0") & ";"
    Next
    s = s & "numfmt=" & PptSync_OptValue(Registry_GetOption(c, "numfmt", "int")) & ";legend=" & PptSync_OptValue(Registry_GetOption(c, "legend", "right")) & ";"
    For Each k In Array("labels", "totals", "connectors", "numfmt", "legend", "cagr", "diff")
        If Kinds_Supports(kind, CStr(k), nSeries) Then caps = caps & k & ","
    Next
    PptSync_OptsText = s & "caps=" & caps
End Function

' Valor de opcao seguro para a etiqueta TQQ_OPTS: so a-z, A-Z e 0-9, no maximo 20 caracteres. Pura.
Public Function PptSync_OptValue(v As String) As String
    Dim i As Long, ch As String, s As String
    For i = 1 To Len(v)
        ch = Mid$(v, i, 1)
        If ch Like "[a-zA-Z0-9]" Then s = s & ch
        If Len(s) >= 20 Then Exit For
    Next
    PptSync_OptValue = s
End Function

' Grava TQQ_KIND e TQQ_OPTS na forma do slide (a cada sincronizacao e colagem). Nunca propaga erro.
Public Sub PptSync_TagOpts(shp As Object, c As Chart)
    On Error Resume Next
    shp.Tags.Add PPTSYNC_TAG_KIND, PptSync_OptValue(Registry_ChartType(c))
    shp.Tags.Add PPTSYNC_TAG_OPTS, PptSync_OptsText(c)
    If Err.Number <> 0 Then Errors_Log "PptSync_TagOpts" & vbTab & Err.Number & vbTab & Err.Description
    Err.Clear
End Sub

' Nome valido para o indice e unico no slide (o indice acha a forma pelo nome): um nome com os separadores do indice ou caracteres que o .pptx
' salvo muda (PptIdx_SafeField) troca eles por "_"; outra forma com o mesmo nome (Shape.Id diferente: objetos de outro processo nao se comparam com Is) faz esta virar
' "<nome>_2", "_3"... O nome da forma do slide nao e usado pelo Excel. Nomes comparados SEM diferenciar maiusculas (medido na C3: Shapes(nome)
' do PowerPoint ignora a caixa; com "tqq_A" e "tqq_a" no slide, os dois nomes achariam a primeira). Igual a PptTqq_EnsureUniqueName do .ppam.
Public Sub PptSync_EnsureUniqueName(shp As Object)
    Dim sl As Object, s As Object, base As String, n As Long, clash As Boolean, myId As Long
    Set sl = shp.Parent
    base = CStr(shp.Name): myId = CLng(shp.Id)
    If PptIdx_SafeField(base) <> base Then
        base = PptIdx_SafeField(base)
        If Not PptSync_SlideHasName(sl, base) Then shp.Name = base: Exit Sub
        clash = True
    End If
    If Not clash Then
        For Each s In sl.Shapes
            If CLng(s.Id) <> myId And StrComp(CStr(s.Name), base, vbTextCompare) = 0 Then clash = True: Exit For
        Next
    End If
    If Not clash Then Exit Sub
    For n = 2 To 999
        If Not PptSync_SlideHasName(sl, base & "_" & n) Then shp.Name = base & "_" & n: Exit Sub
    Next
End Sub

Private Function PptSync_SlideHasName(sl As Object, nm As String) As Boolean
    Dim s As Object
    For Each s In sl.Shapes
        If StrComp(CStr(s.Name), nm, vbTextCompare) = 0 Then PptSync_SlideHasName = True: Exit Function
    Next
End Function

' Liga/desliga "Atualizar automaticamente" de uma forma do ThinQQ e mantem o indice da apresentacao (TQQ_AUTOIDX). Propaga erro do PowerPoint.
' Ordem de proposito: (1) o nome fica definitivo ANTES de ir para o indice (o indice acha a forma pelo nome); (2) a etiqueta da forma (a
' verdade) e gravada antes do indice (o atalho): se a gravacao do indice falhar, a proxima sincronizacao completa o reconstroi a partir das
' etiquetas (PptSync_RebuildIndex).
Public Sub PptSync_SetAuto(shp As Object, turnOn As Boolean)
    Dim sl As Object, pres As Object, idx As String
    Set sl = shp.Parent: Set pres = sl.Parent
    If turnOn Then PptSync_EnsureUniqueName shp
    If turnOn Then shp.Tags.Add PPTSYNC_TAG_AUTO, "1" Else PptSync_TagDelete shp, PPTSYNC_TAG_AUTO
    idx = PptIdx_Set(PptSync_Tag(pres, PPTSYNC_TAG_INDEX), CLng(sl.SlideID), CStr(shp.Name), PptSync_ItemBook(shp), turnOn)
    If Len(idx) = 0 Then PptSync_TagDelete pres, PPTSYNC_TAG_INDEX Else pres.Tags.Add PPTSYNC_TAG_INDEX, idx
End Sub

' Formas do ThinQQ apontadas pelas entradas (formato do indice; a pasta da entrada e ignorada): slide por SlideID, forma pelo nome; a mesma
' forma listada mais de uma vez entra uma vez so. Entrada que nao acha nada e ignorada; PowerPoint ocupado sobe (nunca vira "nao achou").
Public Function PptSync_ResolveEntries(pres As Object, entries As String) As Collection
    Dim out As New Collection, e As Variant, sl As Object, shp As Object, en As Long, seen As Object, k As String
    Set seen = CreateObject("Scripting.Dictionary")
    For Each e In PptIdx_Parse(entries)
        Set sl = Nothing: Set shp = Nothing
        On Error Resume Next
        Set sl = pres.Slides.FindBySlideID(CLng(e(0)))
        If Not sl Is Nothing Then Set shp = sl.Shapes(CStr(e(1)))
        en = Err.Number
        Err.Clear
        On Error GoTo 0
        If Ppt_IsBusyError(en) Then Err.Raise en, "PptSync_ResolveEntries", "PowerPoint ocupado"
        If Not shp Is Nothing Then
            If Len(PptSync_ItemKind(shp)) > 0 Then
                k = PptSync_ShapeKey(CLng(e(0)), shp, CStr(e(1)))
                If Not seen.Exists(k) Then seen.Add k, True: out.Add shp
            End If
        End If
    Next
    Set PptSync_ResolveEntries = out
End Function

' Chave da forma resolvida: SlideID e Shape.Id (o mesmo objeto achado por nomes que o PowerPoint trata como iguais); sem Id legivel, o nome
' da entrada. Nunca propaga erro.
Private Function PptSync_ShapeKey(slideId As Long, shp As Object, entryName As String) As String
    Dim id As Long
    On Error Resume Next
    id = CLng(shp.Id)
    If Err.Number = 0 Then PptSync_ShapeKey = CStr(slideId) & "#" & CStr(id) Else PptSync_ShapeKey = CStr(slideId) & Chr$(31) & entryName
    Err.Clear
End Function

' Reconstroi TQQ_AUTOIDX com as formas automaticas de items (todas as formas do ThinQQ da apresentacao); so grava se mudou. Cada forma e
' isolada (a que falha fica fora do indice e e registrada; as outras seguem). Nunca propaga erro.
Public Sub PptSync_RebuildIndex(pres As Object, items As Collection)
    Dim shp As Object, out As New Collection, s As String, e As Variant
    On Error GoTo Done
    For Each shp In items
        e = PptSync_IndexEntryOf(shp)
        If IsArray(e) Then out.Add e
    Next
    s = PptIdx_Join(out)
    If s <> PptSync_Tag(pres, PPTSYNC_TAG_INDEX) Then
        If Len(s) = 0 Then PptSync_TagDelete pres, PPTSYNC_TAG_INDEX Else pres.Tags.Add PPTSYNC_TAG_INDEX, s
    End If
    Exit Sub
Done:
    Errors_Log "PptSync_RebuildIndex" & vbTab & Err.Number & vbTab & Err.Description
End Sub

' Entrada do indice de uma forma automatica (Array(slideId, nome, pasta)); Empty se a forma nao e automatica ou se a leitura falhou
' (registrado). Nunca propaga erro.
Private Function PptSync_IndexEntryOf(shp As Object) As Variant
    On Error GoTo Fail
    If PptSync_Tag(shp, PPTSYNC_TAG_AUTO) <> "1" Then Exit Function
    PptSync_EnsureUniqueName shp
    PptSync_IndexEntryOf = Array(CLng(shp.Parent.SlideID), CStr(shp.Name), PptSync_ItemBook(shp))
    Exit Function
Fail:
    Errors_Log "PptSync_RebuildIndex: forma fora do indice" & vbTab & Err.Number & vbTab & Err.Description
    PptSync_IndexEntryOf = Empty
End Function

' ---- Sincronizacao da apresentacao inteira (Tarefa A2) ----

' Resumo para o usuario (textos exatos da spec 4.2). missing = nomes dos arquivos ja separados por ", ". Pura.
Public Function PptSync_SummaryText(nOk As Long, nSame As Long, missing As String, nNoChart As Long, nUnlinked As Long, nFail As Long, Optional nEdited As Long = 0, Optional nNoCell As Long = 0) As String
    Dim s As String
    s = "Gráficos do ThinQQ atualizados: " & nOk & "."
    If nSame > 0 Then s = s & " Sem mudança: " & nSame & "."
    If Len(missing) > 0 Then s = s & vbLf & "Pasta de origem não encontrada ou sem acesso (o gráfico ficou como estava): " & missing & ". Se ela foi movida ou renomeada, use no PowerPoint Arquivo > Informações > Editar Links para Arquivos > Alterar Origem."
    If nNoChart > 0 Then s = s & vbLf & "Gráfico não encontrado na pasta de origem (apagado ou recriado): " & nNoChart & "."
    If nUnlinked > 0 Then s = s & vbLf & "Sem vínculo com o Excel (Quebrar Vínculo no PowerPoint): " & nUnlinked & "."
    If nNoCell > 0 Then s = s & vbLf & "Célula do texto vinculado não encontrada na pasta de origem (apagada, ou a pasta não tem vínculos do ThinQQ): " & nNoCell & "."
    If nEdited > 0 Then s = s & vbLf & "Texto editado no PowerPoint fora do campo vinculado (vincule de novo): " & nEdited & "."
    If nFail > 0 Then s = s & vbLf & "Falha ao atualizar: " & nFail & " (veja o log do ThinQQ)."
    PptSync_SummaryText = s
End Function

' A pasta de origem existe para a sincronizacao: aberta neste Excel (FullName igual: vale tambem para o endereco https do OneDrive/SharePoint)
' ou caminho local/de rede existente no disco (mesma regra do Update, PptSync_SourceReachable). Endereco https nao aberto nao tem como ser
' provado: o grafico fica como estava e a pasta e listada no aviso (limite documentado: a pasta tem de estar aberta). Nunca propaga erro.
Public Function PptSync_FileExists(path As String) As Boolean
    PptSync_FileExists = PptSync_SourceReachable(path)
End Function

' Pasta de origem: a ja aberta neste Excel (com as alteracoes nao salvas, como o proprio vinculo le) ou aberta agora SOMENTE LEITURA, sem
' atualizar vinculos, sem MRU e com eventos desligados durante a abertura (registrada em opened para fechar no fim). Nothing se nao abre
' (registrado). Pasta protegida por senha: Open mostraria uma caixa modal INVISIVEL (DisplayAlerts = False nao a evita) e travaria o
' Application.Run; uma Password ficticia faz a senha errada levantar o erro 1004 em vez de perguntar (ignorada em pasta sem senha). A
' recomendacao "somente leitura" e a notificacao de arquivo em uso tambem nao perguntam. SEGURANCA: o caminho vem da apresentacao (de quem a
' fez): a pasta e aberta com as macros DESLIGADAS (AutomationSecurity = 3, msoAutomationSecurityForceDisable, so durante o Open e restaurada
' em todo caminho), entao nenhuma macro dela roda nem ao abrir nem ao fechar.
Private Function PptSync_OpenBook(path As String, opened As Collection) As Workbook
    Dim w As Workbook, prevAlerts As Boolean, prevEvents As Boolean, en As Long, ed As String, prevSec As Long
    For Each w In Application.Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then Set PptSync_OpenBook = w: Exit Function
    Next
    ' so pastas do Excel: .xll (DLL), .iqy, .slk, .htm, .csv... nunca sao abertos por um caminho vindo da apresentacao (revisao final I2)
    If Not PptData_IsWorkbookExt(path) Then Errors_Log "PptSync_OpenBook" & vbTab & "extensao recusada" & vbTab & path: Exit Function
    prevAlerts = Application.DisplayAlerts: prevEvents = Application.EnableEvents
    Application.DisplayAlerts = False: Application.EnableEvents = False
    On Error Resume Next
    prevSec = Application.AutomationSecurity
    Application.AutomationSecurity = 3                          ' msoAutomationSecurityForceDisable: macros da pasta de origem nunca rodam
    Set w = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, Password:="x", IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    If Err.Number <> 0 Then
        en = Err.Number: ed = Err.Description
        Set w = Nothing
        Err.Clear
        Errors_Log "PptSync_OpenBook" & vbTab & en & vbTab & ed & vbTab & path
    End If
    Application.AutomationSecurity = prevSec
    On Error GoTo 0
    Application.DisplayAlerts = prevAlerts: Application.EnableEvents = prevEvents
    If Not w Is Nothing Then opened.Add w
    Set PptSync_OpenBook = w
End Function

' A pasta ja aberta neste Excel com esse FullName (sem diferenciar maiusculas); Nothing se nao esta aberta. Nunca abre nada.
Private Function PptSync_FindOpenBook(path As String) As Workbook
    Dim w As Workbook
    For Each w In Application.Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then Set PptSync_FindOpenBook = w: Exit Function
    Next
End Function

' Fecha sem salvar as pastas abertas pela sincronizacao, com os eventos desligados (restaurados ao fim; nunca propaga erro).
Private Sub PptSync_CloseOpened(opened As Collection)
    Dim w As Variant, prevEvents As Boolean
    On Error Resume Next
    prevEvents = Application.EnableEvents
    If Not (BUILD_CONFIG = "Test" And mTestKeepEventsOnClose) Then Application.EnableEvents = False
    For Each w In opened
        w.Close SaveChanges:=False
    Next
    Application.EnableEvents = prevEvents
End Sub

' O que fazer com uma forma marcada antes de olhar a pasta de origem (pura): linkState de Ppt_LinkState (1 vinculado, 0 sem vinculo, -1
' ilegivel) e o caminho de origem ja resolvido (PptSync_SourcePath; "" = nem SourceFullName nem TQQ_BOOK). Nunca sincroniza sem estado 1.
Public Function PptSync_Classify(linkState As Long, path As String) As Long
    If linkState = 0 Then
        PptSync_Classify = PPTSYNC_C_UNLINKED
    ElseIf linkState <> 1 Then
        PptSync_Classify = PPTSYNC_C_UNKNOWN
    ElseIf Len(Trim$(path)) = 0 Then
        PptSync_Classify = PPTSYNC_C_NOPATH
    Else
        PptSync_Classify = PPTSYNC_C_PROCEED
    End If
End Function

' Nome do arquivo para o aviso: depois da ultima barra (invertida ou normal: endereco https do OneDrive/SharePoint); "" = origem desconhecida.
Private Function PptSync_DisplayName(path As String) As String
    Dim p As String, k As Long
    p = Trim$(path)
    If Len(p) = 0 Then PptSync_DisplayName = PPTSYNC_UNKNOWN_SRC: Exit Function
    k = InStrRev(p, "\")
    If InStrRev(p, "/") > k Then k = InStrRev(p, "/")
    PptSync_DisplayName = Mid$(p, k + 1)
End Function

' Lista do aviso "Pasta de origem nao encontrada ou sem acesso" (", "): o nome do arquivo de cada caminho (paths = Variant com os caminhos,
' sem repeticao); caminhos diferentes com o MESMO nome aparecem pelo caminho completo para nao se confundirem. Pura.
Public Function PptSync_MissingText(paths As Variant) As String
    Dim i As Long, j As Long, lo As Long, hi As Long, nm() As String, dup As Boolean, s As String, lbl As String
    lo = LBound(paths): hi = UBound(paths)
    If hi < lo Then Exit Function
    ReDim nm(lo To hi)
    For i = lo To hi
        nm(i) = PptSync_DisplayName(CStr(paths(i)))
    Next
    For i = lo To hi
        dup = False
        For j = lo To hi
            If j <> i Then
                If StrComp(nm(i), nm(j), vbTextCompare) = 0 Then dup = True: Exit For
            End If
        Next
        If dup And Len(Trim$(CStr(paths(i)))) > 0 Then lbl = Trim$(CStr(paths(i))) Else lbl = nm(i)
        s = s & IIf(Len(s) > 0, ", ", "") & lbl
    Next
    PptSync_MissingText = s
End Function

' Sincroniza UMA forma ja ligada a pasta de origem wb, isolada: qualquer erro (achar o grafico, impressao digital, SyncShape) conta a forma como
' falha, registra e devolve PPTSYNC_R_FAIL sem levantar, EXCETO PowerPoint ocupado (Ppt_IsBusyError), que sobe para quem chama poder tentar
' de novo (PptSync_Items fecha as pastas que abriu e repassa). A TQQ_FP e calculada DEPOIS da sincronizacao (descreve o que o slide acabou de receber,
' igual a colagem); se so a gravacao dela falha, o grafico conta como atualizado, a TQQ_FP velha e apagada (senao uma volta do dado ao valor
' antigo seria pulada como "sem mudanca") e fica registrado.
Private Function PptSync_SyncOne(shp As Object, wb As Workbook, force As Boolean) As Long
    Dim co As ChartObject, fp As String, skipped As Boolean, bn As Long, bd As String
    On Error GoTo Fail
    PptSync_Hook "findchart"
    Set co = Registry_FindChartObject(wb, PptSync_Tag(shp, PPTSYNC_TAG_CHART))
    If co Is Nothing Then PptSync_SyncOne = PPTSYNC_R_NOCHART: Exit Function
    PptSync_Hook "fingerprint"
    fp = PptSync_Fingerprint(co, shp)
    If Not force And fp = PptSync_Tag(shp, PPTSYNC_TAG_FP) Then PptSync_SyncOne = PPTSYNC_R_SAME: Exit Function
    PptSync_SyncShape shp, co, True, skipped
    If skipped Then PptSync_SyncOne = PPTSYNC_R_SKIPPED: Exit Function   ' Update ignorado: sem TQQ_FP, sem contar como atualizado
    PptSync_SyncOne = PPTSYNC_R_OK
    PptSync_TagOpts shp, co.Chart                      ' estado das opcoes para a faixa do PowerPoint (nunca propaga erro)
    On Error GoTo FpFail
    PptSync_Hook "tagfp"
    shp.Tags.Add PPTSYNC_TAG_FP, PptSync_Fingerprint(co, shp)
    Exit Function
FpFail:
    Errors_Log "PptSync/TQQ_FP" & vbTab & Err.Number & vbTab & Err.Description
    On Error GoTo -1
    On Error Resume Next
    shp.Tags.Delete PPTSYNC_TAG_FP
    Exit Function
Fail:
    If Ppt_IsBusyError(Err.Number) Then
        bn = Err.Number: bd = Err.Description
        On Error GoTo -1
        On Error GoTo 0                                 ' sem isto o GoTo -1 rearma o Fail e o Raise voltaria para ca (laco)
        Err.Raise bn, "PptSync_SyncOne", bd
    End If
    Errors_Log "PptSync" & vbTab & Err.Number & vbTab & Err.Description
    PptSync_SyncOne = PPTSYNC_R_FAIL
End Function

' Sincroniza todos os elementos do ThinQQ (formas soltas marcadas) da apresentacao e reconstroi o indice das automaticas. force = False pula
' os que nao mudaram (TQQ_FP). Regras de origem, aviso e erros: PptSync_Items. Devolve o resumo (PptSync_SummaryText).
Public Function PptSync_Presentation(pres As Object, force As Boolean) As String
    Dim sl As Object, shp As Object, items As Collection
    PptSync_Hook "presentation"
    Set items = New Collection
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If Len(PptSync_ItemKind(shp)) > 0 Then items.Add shp
        Next
    Next
    If items.Count = 0 Then
        If Len(PptSync_Tag(pres, PPTSYNC_TAG_INDEX)) > 0 Then PptSync_TagDelete pres, PPTSYNC_TAG_INDEX    ' indice velho de formas apagadas
        PptSync_Presentation = PPTSYNC_NONE
        Exit Function
    End If
    PptSync_RebuildIndex pres, items
    PptSync_Presentation = PptSync_Items(items, force)
End Function

' Sincroniza as formas dadas (todas da mesma apresentacao ou nao). Pasta de origem inexistente ou que nao abre (senha, corrompida): o elemento
' fica como esta e a pasta entra no aviso (NUNCA LinkFormat.Update com a origem ausente: travaria o PowerPoint). Update ignorado por origem
' inalcancavel: nao conta como atualizado e a TQQ_FP nao e gravada. Forma sem vinculo ou com estado ilegivel nunca e sincronizada. Pastas
' abertas por esta chamada sao fechadas sem salvar, com a tela congelada (restaurada em todo caminho). Erro de um elemento e contado
' (PptSync_SyncOne); PowerPoint ocupado e falha estrutural sobem. noOpen (ciclo automatico): uma pasta que nao esta aberta neste Excel nunca
' e aberta; entra no aviso de pasta nao encontrada.
Public Function PptSync_Items(items As Collection, force As Boolean, Optional noOpen As Boolean = False) As String
    Dim shp As Object, opened As Collection, books As Object, missingSet As Object, path As String, k As Long
    Dim wb As Workbook, missing As String, en As Long, ed As String, prevSU As Boolean, ls As Long
    Dim nOk As Long, nSame As Long, nNoChart As Long, nUnlinked As Long, nFail As Long, kind As String, r As Long, nEdited As Long, nNoCell As Long
    Set opened = New Collection
    Set books = CreateObject("Scripting.Dictionary")
    books.CompareMode = 1
    Set missingSet = CreateObject("Scripting.Dictionary")
    missingSet.CompareMode = 1
    prevSU = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False
    For Each shp In items
        k = k + 1
        Application.StatusBar = "ThinQQ: atualizando elemento " & k & " de " & items.Count & " do PowerPoint..."
        kind = PptSync_ItemKind(shp)
        If kind = "chart" Then ls = Ppt_LinkState(shp) Else ls = 1  ' texto vinculado: sem vinculo nativo (a pasta vem de TQQ_BOOK)
        path = ""
        If ls = 1 Then path = PptSync_ItemBook(shp)
        Select Case PptSync_Classify(ls, path)
            Case PPTSYNC_C_UNLINKED
                nUnlinked = nUnlinked + 1
            Case PPTSYNC_C_UNKNOWN
                nFail = nFail + 1
                Errors_Log "PptSync: estado do vinculo ilegivel; forma nao sincronizada" & vbTab & PptSync_Tag(shp, PPTSYNC_TAG_CHART)
            Case PPTSYNC_C_NOPATH
                If Not missingSet.Exists("") Then missingSet.Add "", True
            Case Else
                If Not books.Exists(path) Then
                    Set wb = Nothing
                    If noOpen Then
                        Set wb = PptSync_FindOpenBook(path)
                    ElseIf PptSync_FileExists(path) Then
                        Set wb = PptSync_OpenBook(path, opened)
                    End If
                    If wb Is Nothing Then missingSet(path) = True
                    books.Add path, wb
                End If
                Set wb = books.Item(path)
                If Not wb Is Nothing Then
                    If kind = "text" Then r = PptText_Sync(shp, wb, force) Else r = PptSync_SyncOne(shp, wb, force)
                    Select Case r
                        Case PPTSYNC_R_OK: nOk = nOk + 1
                        Case PPTSYNC_R_SAME: nSame = nSame + 1
                        Case PPTSYNC_R_NOCHART
                            If kind = "text" Then nNoCell = nNoCell + 1 Else nNoChart = nNoChart + 1
                        Case PPTSYNC_R_SKIPPED: missingSet(path) = True
                        Case PPTSYNC_R_EDITED: nEdited = nEdited + 1
                        Case Else: nFail = nFail + 1
                    End Select
                End If
        End Select
    Next
    PptSync_CloseOpened opened
    Application.ScreenUpdating = prevSU
    Application.StatusBar = False
    If missingSet.Count > 0 Then missing = PptSync_MissingText(missingSet.Keys)
    PptSync_Items = PptSync_SummaryText(nOk, nSame, missing, nNoChart, nUnlinked, nFail, nEdited, nNoCell)
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptSync_CloseOpened opened
    Application.ScreenUpdating = prevSU
    Application.StatusBar = False
    On Error GoTo 0
    Err.Raise en, "PptSync_Items", ed
End Function
