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
Public Const CORE_PPT_API As Long = 1                       ' versao da API do core usada pelo suplemento do PowerPoint (Core_PptApiVersion)
Public Const PPTSYNC_NONE As String = "Esta apresentação não tem gráficos do ThinQQ vinculados (copiados com 'Copiar como gráfico vinculado')."
Public Const PPTSYNC_NO_PPT As String = "Abra a apresentação no PowerPoint e use de novo 'Atualizar gráficos no PowerPoint'."
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
Private mUpdateCount As Long                                ' chamadas de LinkFormat.Update feitas por PptSync_SyncShape (lido so no build Test)

' Gancho de teste: o proximo ponto nomeado ("update", "snapshot", "apply") levanta erro (uso unico). Inerte fora do build Test.
Public Sub PptSync_SetFailForTest(point As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestFail = point
End Sub

Public Sub PptSync_ResetTestHooks()
    mTestFail = ""
    mUpdateCount = 0
End Sub

' Gancho de teste: quantas vezes PptSync_SyncShape chamou LinkFormat.Update desde o ultimo PptSync_ResetTestHooks (0 fora do build Test).
Public Function PptSync_UpdateCountForTest() As Long
    If BUILD_CONFIG <> "Test" Then Exit Function
    PptSync_UpdateCountForTest = mUpdateCount
End Function

Private Sub PptSync_Hook(point As String)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    If mTestFail <> point Then Exit Sub
    mTestFail = ""
    Err.Raise vbObjectError + 9992, "modPptSync", "falha injetada (" & point & ")"
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
    Dim wb As Workbook, p As String
    On Error GoTo Done
    p = Trim$(path)
    If Len(p) = 0 Then Exit Function
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, p, vbTextCompare) = 0 Then PptSync_SourceReachable = True: Exit Function
    Next
    If LCase$(Left$(p, 7)) = "http://" Or LCase$(Left$(p, 8)) = "https://" Then Exit Function
    PptSync_SourceReachable = CBool(CreateObject("Scripting.FileSystemObject").FileExists(p))
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
    b = StrConv(s, vbFromUnicode)
    PptSync_Fingerprint = Sha256_Bytes(b, UBound(b) - LBound(b) + 1)
End Function

' ---- Sincronizacao da apresentacao inteira (Tarefa A2) ----

' Resumo para o usuario (textos exatos da spec 4.2). missing = nomes dos arquivos ja separados por ", ". Pura.
Public Function PptSync_SummaryText(nOk As Long, nSame As Long, missing As String, nNoChart As Long, nUnlinked As Long, nFail As Long) As String
    Dim s As String
    s = "Gráficos do ThinQQ atualizados: " & nOk & "."
    If nSame > 0 Then s = s & " Sem mudança: " & nSame & "."
    If Len(missing) > 0 Then s = s & vbLf & "Pasta de origem não encontrada ou sem acesso (o gráfico ficou como estava): " & missing & ". Se ela foi movida ou renomeada, use no PowerPoint Arquivo > Informações > Editar Links para Arquivos > Alterar Origem."
    If nNoChart > 0 Then s = s & vbLf & "Gráfico não encontrado na pasta de origem (apagado ou recriado): " & nNoChart & "."
    If nUnlinked > 0 Then s = s & vbLf & "Sem vínculo com o Excel (Quebrar Vínculo no PowerPoint): " & nUnlinked & "."
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
' atualizar vinculos, sem MRU e com eventos desligados durante a abertura (registrada em opened para fechar no fim). Nothing se nao abre.
Private Function PptSync_OpenBook(path As String, opened As Collection) As Workbook
    Dim w As Workbook, prevAlerts As Boolean, prevEvents As Boolean
    For Each w In Application.Workbooks
        If StrComp(w.FullName, path, vbTextCompare) = 0 Then Set PptSync_OpenBook = w: Exit Function
    Next
    prevAlerts = Application.DisplayAlerts: prevEvents = Application.EnableEvents
    Application.DisplayAlerts = False: Application.EnableEvents = False
    On Error Resume Next
    Set w = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    If Err.Number <> 0 Then Errors_Log "PptSync_OpenBook" & vbTab & Err.Number & vbTab & Err.Description: Set w = Nothing
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = prevAlerts: Application.EnableEvents = prevEvents
    If Not w Is Nothing Then opened.Add w
    Set PptSync_OpenBook = w
End Function

Private Sub PptSync_CloseOpened(opened As Collection)
    Dim w As Variant
    On Error Resume Next
    For Each w In opened
        w.Close SaveChanges:=False
    Next
End Sub

' True se PptSync_SyncShape terminou sem erro. updateSkipped = True quando o LinkFormat.Update foi ignorado por origem inalcancavel.
Private Function PptSync_TrySync(shp As Object, co As ChartObject, ByRef updateSkipped As Boolean) As Boolean
    On Error GoTo Fail
    PptSync_SyncShape shp, co, True, updateSkipped
    PptSync_TrySync = True
    Exit Function
Fail:
    Errors_Log "PptSync" & vbTab & Err.Number & vbTab & Err.Description
End Function

' Sincroniza todos os graficos ThinQQ vinculados (formas soltas com a etiqueta TQQ_CHART) da apresentacao. force = False pula os que nao
' mudaram (TQQ_FP). Pasta de origem inexistente ou que nao abre: o grafico fica como esta (NUNCA LinkFormat.Update: travaria o PowerPoint).
' Se o Update foi ignorado por origem inalcancavel (a pasta sumiu depois da conferencia), o grafico NAO conta como atualizado, a TQQ_FP NAO e
' gravada e a pasta entra no aviso. Pastas abertas por esta chamada sao fechadas sem salvar. Devolve o resumo (PptSync_SummaryText).
' Erro de um grafico e contado e registrado; so falha estrutural sobe.
Public Function PptSync_Presentation(pres As Object, force As Boolean) As String
    Dim sl As Object, shp As Object, items As Collection, opened As Collection, books As Object, path As String, k As Long
    Dim wb As Workbook, co As ChartObject, fp As String, missing As String, en As Long, ed As String, skipped As Boolean
    Dim nOk As Long, nSame As Long, nNoChart As Long, nUnlinked As Long, nFail As Long
    Set items = New Collection: Set opened = New Collection
    Set books = CreateObject("Scripting.Dictionary")
    books.CompareMode = 1                                       ' texto: caminhos sem diferenciar maiusculas
    For Each sl In pres.Slides
        For Each shp In sl.Shapes
            If Len(PptSync_Tag(shp, PPTSYNC_TAG_CHART)) > 0 Then items.Add shp
        Next
    Next
    If items.Count = 0 Then PptSync_Presentation = PPTSYNC_NONE: Exit Function
    On Error GoTo Fail
    For Each shp In items
        k = k + 1
        Application.StatusBar = "ThinQQ: atualizando gráfico " & k & " de " & items.Count & " do PowerPoint..."
        If Ppt_LinkState(shp) = 0 Then
            nUnlinked = nUnlinked + 1
        Else
            path = PptSync_SourcePath(PptSync_LinkSource(shp), PptSync_Tag(shp, PPTSYNC_TAG_BOOK))
            If Not books.Exists(path) Then
                Set wb = Nothing
                If PptSync_FileExists(path) Then Set wb = PptSync_OpenBook(path, opened)
                If wb Is Nothing Then missing = PptSync_AddMissing(missing, path)
                books.Add path, wb
            End If
            Set wb = books.Item(path)
            If Not wb Is Nothing Then
                Set co = Registry_FindChartObject(wb, PptSync_Tag(shp, PPTSYNC_TAG_CHART))
                If co Is Nothing Then
                    nNoChart = nNoChart + 1
                Else
                    fp = PptSync_Fingerprint(co, shp)
                    If Not force And fp = PptSync_Tag(shp, PPTSYNC_TAG_FP) Then
                        nSame = nSame + 1
                    ElseIf Not PptSync_TrySync(shp, co, skipped) Then
                        nFail = nFail + 1
                    ElseIf skipped Then
                        missing = PptSync_AddMissing(missing, path)     ' Update ignorado: sem TQQ_FP, sem contar como atualizado
                    Else
                        shp.Tags.Add PPTSYNC_TAG_FP, fp
                        nOk = nOk + 1
                    End If
                End If
            End If
        End If
    Next
    PptSync_CloseOpened opened
    Application.StatusBar = False
    PptSync_Presentation = PptSync_SummaryText(nOk, nSame, missing, nNoChart, nUnlinked, nFail)
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    PptSync_CloseOpened opened
    Application.StatusBar = False
    On Error GoTo 0
    Err.Raise en, "PptSync_Presentation", ed
End Function

' Acrescenta o nome do arquivo de path a lista do aviso (separada por ", ") sem repetir.
Private Function PptSync_AddMissing(missing As String, path As String) As String
    Dim nm As String
    nm = Util_FileName(path)
    If InStr(", " & missing & ",", ", " & nm & ",") > 0 Then PptSync_AddMissing = missing: Exit Function
    PptSync_AddMissing = missing & IIf(Len(missing) > 0, ", ", "") & nm
End Function
