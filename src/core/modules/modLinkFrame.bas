Attribute VB_Name = "modLinkFrame"
Option Explicit
' Moldura do intervalo de origem de um grafico do ThinQQ que foi ao PowerPoint (como a "data link frame" verde do think-cell; spec 2026-10-15
' secao 5.2): retangulo sem preenchimento, linha verde de 2 pt, que acompanha linhas/colunas inseridas (Placement = xlMoveAndSize, medido),
' nao impresso e bloqueado. Nome tqqlink_<id do grafico> (fora de ID_PREFIX + id: nao e ChartObject). Arrastar a moldura NAO muda o vinculo.
' Alterar dados recria o grafico com id novo e o slide deixa de acompanha-lo (o grafico do slide le as celulas antigas: sonda P-U6, 1.13.0): a
' moldura sai (nao e redesenhada) e o usuario e avisado para inserir ou colar de novo (LINKFRAME_NOTE_REBUILT). Cor pelo estado (1.13.0,
' LinkFrame_SetState): verde = alguma apresentacao aberta usa o grafico; vermelha = nenhuma. Nada aqui propaga erro.
Public Const LINKFRAME_PREFIX As String = "tqqlink_"
Public Const LINKFRAME_NOTE_REBUILT As String = "O gráfico foi recriado: o gráfico que já está no slide do PowerPoint não o acompanha mais e, ao atualizar o vínculo, pode perder os totais e mostrar dados antigos. Cole o gráfico de novo no slide (Mais > Copiar como gráfico vinculado) e apague o antigo."
Public Const LINKFRAME_MSG_NONE As String = "Esta pasta não tem molduras de intervalos vinculados ao PowerPoint (elas aparecem ao usar Vincular ao PowerPoint ou Copiar como gráfico vinculado)."
Public Const LINKFRAME_GREEN As Long = 5287936      ' RGB(0, 176, 80)
Public Const LINKFRAME_RED As Long = 192            ' RGB(192, 0, 0)
Private mTestFailDraw As Boolean    ' gancho de teste (LinkFrame_SetFailDrawForTest): falha logo depois de criar o retangulo; so no build Test

Public Function LinkFrame_Name(cht As Chart) As String
    LinkFrame_Name = LINKFRAME_PREFIX & Registry_IdFromChart(cht)
End Function

' Gancho de teste: a proxima LinkFrame_Draw falha depois do AddShape (o retangulo tem de sumir). Inerte fora do build Test; zerado no uso.
Public Sub LinkFrame_SetFailDrawForTest(v As Boolean)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestFailDraw = v
End Sub

' Desenho isolado (sonda pendente P4: se o retangulo atrapalhar cliques nas celulas, trocar por 4 linhas AddLine aqui). O retangulo recebe o
' nome logo depois do AddShape e, em qualquer falha, e apagado antes de re-levantar: nunca sobra um retangulo azul sem nome tqqlink_.
Private Function LinkFrame_Draw(ws As Worksheet, rng As Range, nm As String) As Shape
    Dim s As Shape, en As Long, ed As String, failNow As Boolean
    failNow = mTestFailDraw
    mTestFailDraw = False
    Set s = ws.Shapes.AddShape(msoShapeRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    On Error GoTo Fail
    s.Name = nm
    If BUILD_CONFIG = "Test" Then
        If failNow Then Err.Raise vbObjectError + 7300, "LinkFrame_Draw", "falha injetada (teste)"
    End If
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = LINKFRAME_GREEN
    s.Line.Weight = 2
    s.Placement = xlMoveAndSize
    s.Locked = True
    On Error Resume Next
    s.DrawingObject.PrintObject = False
    Err.Clear
    Set LinkFrame_Draw = s
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description
    On Error GoTo -1
    On Error Resume Next
    s.Delete
    On Error GoTo 0
    Err.Raise en, "LinkFrame_Draw", ed
End Function

' A moldura do grafico (em qualquer planilha da pasta do grafico) ou Nothing.
Public Function LinkFrame_Find(cht As Chart) As Shape
    Dim ws As Worksheet, s As Shape, nm As String
    On Error Resume Next
    nm = LinkFrame_Name(cht)
    For Each ws In Registry_Workbook(cht).Worksheets
        Set s = Nothing
        Set s = ws.Shapes(nm)
        Err.Clear
        If Not s Is Nothing Then Set LinkFrame_Find = s: Exit Function
    Next
End Function

' Cor da moldura pelo estado (como o think-cell): verde = alguma apresentacao aberta usa o grafico; vermelha = nenhuma. So grava se mudou
' (mudar a cor altera a pasta). Nunca propaga.
Public Sub LinkFrame_SetState(cht As Chart, used As Boolean)
    Dim s As Shape
    On Error Resume Next
    Set s = LinkFrame_Find(cht)
    If s Is Nothing Then Exit Sub
    LinkFrame_SetShapeState s, used
    Err.Clear
End Sub

' A mesma cor de estado numa moldura ja achada (Core_RefreshFrames pinta cada forma, inclusive a de uma planilha copiada). Nunca propaga.
Public Sub LinkFrame_SetShapeState(s As Shape, used As Boolean)
    Dim want As Long
    On Error Resume Next
    If s Is Nothing Then Exit Sub
    If used Then want = LINKFRAME_GREEN Else want = LINKFRAME_RED
    If s.Line.ForeColor.RGB <> want Then s.Line.ForeColor.RGB = want
    Err.Clear
End Sub

Public Sub LinkFrame_Remove(cht As Chart)
    Dim s As Shape, guard As Long
    On Error Resume Next
    Do
        Set s = LinkFrame_Find(cht)
        If s Is Nothing Then Exit Do
        s.Delete
        guard = guard + 1
    Loop While guard < 20
End Sub

' Desenha (ou redesenha) a moldura em volta de rng. Devolve a forma (Nothing se falhou: so registrado).
Public Function LinkFrame_Ensure(cht As Chart, rng As Range) As Shape
    Dim s As Shape
    On Error GoTo Fail
    LinkFrame_Remove cht
    Set s = LinkFrame_Draw(rng.Worksheet, rng, LinkFrame_Name(cht))
    s.AlternativeText = "ThinQQ: intervalo do gráfico " & Registry_IdFromChart(cht) & " vinculado ao PowerPoint"
    Set LinkFrame_Ensure = s
    Exit Function
Fail:
    Errors_Log "LinkFrame_Ensure" & vbTab & Err.Number & vbTab & Err.Description
End Function

' Molduras cujo grafico nao existe mais na pasta (todas as planilhas: Core_Insert, Core_Refresh, Alterar dados).
Public Sub LinkFrame_CleanOrphans(wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        LinkFrame_CleanOrphansSheet ws
    Next
    Err.Clear
End Sub

' So as molduras orfas de uma planilha (Watchers_Rehook: a planilha ativada; o grafico e procurado na pasta inteira).
Public Sub LinkFrame_CleanOrphansSheet(ws As Worksheet)
    Dim wb As Workbook, i As Long, nm As String
    On Error Resume Next
    Set wb = ws.Parent
    For i = ws.Shapes.Count To 1 Step -1
        nm = ws.Shapes(i).Name
        If Left$(nm, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then
            If Registry_FindChartObject(wb, Mid$(nm, Len(LINKFRAME_PREFIX) + 1)) Is Nothing Then ws.Shapes(i).Delete
        End If
    Next
    Err.Clear
End Sub

' Mostrar/ocultar todas as molduras da pasta: alguma visivel => oculta todas (0); todas ocultas => mostra (1); nenhuma moldura => -1.
Public Function LinkFrame_ToggleAll(wb As Workbook) As Long
    Dim ws As Worksheet, s As Shape, anyVis As Boolean, n As Long
    On Error Resume Next
    For Each ws In wb.Worksheets
        For Each s In ws.Shapes
            If Left$(s.Name, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then
                n = n + 1
                If s.Visible = msoTrue Then anyVis = True
            End If
        Next
    Next
    If n = 0 Then LinkFrame_ToggleAll = -1: Exit Function
    For Each ws In wb.Worksheets
        For Each s In ws.Shapes
            If Left$(s.Name, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then s.Visible = IIf(anyVis, msoFalse, msoTrue)
        Next
    Next
    LinkFrame_ToggleAll = IIf(anyVis, 0, 1)
End Function
