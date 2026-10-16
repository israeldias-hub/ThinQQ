Attribute VB_Name = "modLinkFrame"
Option Explicit
' Moldura do intervalo de origem de um grafico do ThinQQ que foi ao PowerPoint (como a "data link frame" verde do think-cell; spec 2026-10-15
' secao 5.2): retangulo sem preenchimento, linha verde de 2 pt, que acompanha linhas/colunas inseridas (Placement = xlMoveAndSize, medido),
' nao impresso e bloqueado. Nome tqqlink_<id do grafico> (fora de ID_PREFIX + id: nao e ChartObject). Arrastar a moldura NAO muda o vinculo
' (mudar o intervalo = Alterar dados, que redesenha a moldura). Nada aqui propaga erro.
Public Const LINKFRAME_PREFIX As String = "tqqlink_"

Public Function LinkFrame_Name(cht As Chart) As String
    LinkFrame_Name = LINKFRAME_PREFIX & Registry_IdFromChart(cht)
End Function

' Desenho isolado (sonda pendente P4: se o retangulo atrapalhar cliques nas celulas, trocar por 4 linhas AddLine aqui).
Private Function LinkFrame_Draw(ws As Worksheet, rng As Range) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = RGB(0, 176, 80)
    s.Line.Weight = 2
    s.Placement = xlMoveAndSize
    s.Locked = True
    On Error Resume Next
    s.DrawingObject.PrintObject = False
    Err.Clear
    Set LinkFrame_Draw = s
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
    Set s = LinkFrame_Draw(rng.Worksheet, rng)
    s.Name = LinkFrame_Name(cht)
    s.AlternativeText = "ThinQQ: intervalo do gráfico " & Registry_IdFromChart(cht) & " vinculado ao PowerPoint"
    Set LinkFrame_Ensure = s
    Exit Function
Fail:
    Errors_Log "LinkFrame_Ensure" & vbTab & Err.Number & vbTab & Err.Description
End Function

' Molduras cujo grafico nao existe mais na pasta.
Public Sub LinkFrame_CleanOrphans(wb As Workbook)
    Dim ws As Worksheet, i As Long, nm As String
    On Error Resume Next
    For Each ws In wb.Worksheets
        For i = ws.Shapes.Count To 1 Step -1
            nm = ws.Shapes(i).Name
            If Left$(nm, Len(LINKFRAME_PREFIX)) = LINKFRAME_PREFIX Then
                If Registry_FindChartObject(wb, Mid$(nm, Len(LINKFRAME_PREFIX) + 1)) Is Nothing Then ws.Shapes(i).Delete
            End If
        Next
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
