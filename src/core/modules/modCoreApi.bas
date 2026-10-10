Attribute VB_Name = "modCoreApi"
Option Explicit
' API pública chamada pelo loader via Application.Run. Nunca propaga erro.
Private mLoaderPath As String

Public Function Core_Ping() As String
    Core_Ping = "pong"
End Function

Public Function Core_Version() As String
    Core_Version = APP_VERSION
End Function

Public Function Core_LoaderPath() As String
    Core_LoaderPath = mLoaderPath
End Function

Public Sub Core_OnLoaded()
    On Error Resume Next
    Watchers_Start
End Sub

Public Sub Core_OnUnload()
    On Error Resume Next
    Watchers_Stop
End Sub

Public Sub Core_RegisterLoader(loaderName As String, loaderFullPath As String)
    Watchers_SetLoader loaderName
    mLoaderPath = loaderFullPath
End Sub

' Mensagem ao usuário (suprimida em build de teste para não travar o Excel invisível).
Public Sub Core_Notify(msg As String, Optional style As VbMsgBoxStyle = vbInformation)
    Errors_Log "notify" & vbTab & Replace(msg, vbLf, " | ")
    If BUILD_CONFIG = "Test" Then Exit Sub
    MsgBox msg, style, APP_NAME
End Sub

Public Function Core_ActiveChart() As Chart
    Dim c As Chart
    On Error Resume Next
    Set c = Application.ActiveChart
    On Error GoTo 0
    If c Is Nothing Then Exit Function
    If Registry_IsThinQQChart(c) Then Set Core_ActiveChart = c
End Function

Public Function Core_HasChart() As Boolean
    Core_HasChart = Not Core_ActiveChart() Is Nothing
End Function

Public Sub Core_Insert(kind As String)
    On Error GoTo Fail
    Dim data As clsChartData, msg As String, co As ChartObject, ws As Worksheet
    If TypeName(Selection) <> "Range" Then
        Core_Notify "Selecione a tabela de dados (cabeçalho + categorias) antes de inserir o gráfico."
        Exit Sub
    End If
    Set data = ChartData_FromRange(Selection)
    msg = data.Validate()
    If Len(msg) > 0 Then Core_Notify msg, vbExclamation: Exit Sub
    Set ws = data.Source.Worksheet
    Application.ScreenUpdating = False
    CalcSheet_Cleanup ws.Parent
    If kind = "waterfall" Then
        Set co = Waterfall_Insert(ws, data)
    Else
        Set co = Stacked_Insert(ws, data, kind)
    End If
    Application.ScreenUpdating = True
    Watchers_Rehook
    co.Select
    Core_SafeRefresh co.Chart
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Errors_Handle "Core_Insert(" & kind & ")"
End Sub

Public Function Core_GetToggle(key As String) As Boolean
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Function
    Core_GetToggle = (Registry_GetOption(c, key, "0") = "1")
End Function

' Aplica o toggle no gráfico informado (sem depender da seleção). Propaga erros.
Private Sub ApplyToggle(c As Chart, key As String, value As Boolean)
    Dim kind As String
    kind = Registry_ChartType(c)
    Select Case key
        Case "totals"
            If kind <> "waterfall" Then Stacked_SetTotals c, value
        Case "labels"
            If kind = "waterfall" Then Waterfall_SetLabels c, value Else Stacked_SetLabels c, value
        Case "connectors"
            If kind = "waterfall" Then Ann_SetWaterfallConnectors c, value Else Stacked_SetConnectors c, value
    End Select
End Sub

' Aplica a opção no gráfico informado (sem depender da seleção). Propaga erros.
Private Sub ApplyOption(c As Chart, key As String, value As String)
    Registry_SetOption c, key, value
    Select Case key
        Case "numfmt"
            Style_ApplyNumberFormat c, Style_NumberFormat(value)
        Case "palette"
            If Registry_ChartType(c) = "waterfall" Then Waterfall_ApplyColors c, value Else Style_ApplySeriesColors c, value
        Case "gap"
            On Error Resume Next
            c.ChartGroups(1).GapWidth = CLng(value)
            On Error GoTo 0
    End Select
End Sub

Public Sub Core_SetToggle(key As String, value As Boolean)
    On Error GoTo Fail
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    ApplyToggle c, key, value
    Core_SafeRefresh c
    Exit Sub
Fail:
    Errors_Handle "Core_SetToggle(" & key & ")"
End Sub

Public Function Core_GetOption(key As String) As String
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Function
    Core_GetOption = Registry_GetOption(c, key, "")
End Function

Public Sub Core_SetOption(key As String, value As String)
    On Error GoTo Fail
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    ApplyOption c, key, value
    Core_SafeRefresh c
    Exit Sub
Fail:
    Errors_Handle "Core_SetOption(" & key & ")"
End Sub

Public Sub Core_AddAnnotation(kind As String, Optional fromIdx As Long = 0, Optional toIdx As Long = 0)
    On Error GoTo Fail
    Dim c As Chart, n As Long
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    n = Geom_ChartCategoryCount(c)
    If n < 2 Then Core_Notify "O gráfico precisa de pelo menos 2 categorias.": Exit Sub
    If fromIdx < 1 Or toIdx > n Or fromIdx >= toIdx Then
        If Not Core_AskCategories(c, n, fromIdx, toIdx) Then Exit Sub
    End If
    If kind = "cagr" Then Ann_AddCagr c, fromIdx, toIdx Else Ann_AddDifference c, fromIdx, toIdx
    Exit Sub
Fail:
    If Err.Number = vbObjectError + 7001 Then
        ' Seta de diferença em gráfico 100%: aviso ao usuário, não é falha.
        Core_Notify Err.Description, vbInformation
    Else
        Errors_Handle "Core_AddAnnotation(" & kind & ")"
    End If
End Sub

Private Function Core_AskCategories(c As Chart, n As Long, ByRef fromIdx As Long, ByRef toIdx As Long) As Boolean
    Dim frm As frmAnnotation, cats() As String, body As Range, i As Long
    Set body = Block_Body(Registry_Block(c))
    ReDim cats(1 To n)
    For i = 1 To n
        cats(i) = CStr(body.Cells(i + 1, 1).Text)
    Next
    Set frm = New frmAnnotation
    frm.Prepare cats, 1, n
    frm.Show
    If frm.Result = "ok" Then
        fromIdx = frm.FromIdx: toIdx = frm.ToIdx
        Core_AskCategories = True
    End If
    Unload frm
End Function

Public Sub Core_RemoveAnnotations()
    On Error GoTo Fail
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    Ann_RemoveAll c, False
    Exit Sub
Fail:
    Errors_Handle "Core_RemoveAnnotations"
End Sub

Public Sub Core_SafeRefresh(cht As Chart)
    On Error Resume Next
    Labels_RefreshPositions cht
    Ann_Refresh cht
End Sub

' Re-layout do gráfico ativo (ou de todos da planilha) e limpeza de blocos órfãos.
Public Sub Core_Refresh()
    On Error GoTo Fail
    Dim c As Chart, co As ChartObject, ws As Worksheet
    Set c = Core_ActiveChart()
    If Not c Is Nothing Then
        Core_SafeRefresh c
        CalcSheet_Cleanup Registry_Workbook(c)
    Else
        Set ws = ActiveSheet
        For Each co In ws.ChartObjects
            If Left$(co.Name, Len(ID_PREFIX)) = ID_PREFIX Then Core_SafeRefresh co.Chart
        Next
        CalcSheet_Cleanup ws.Parent
    End If
    Watchers_Rehook
    Exit Sub
Fail:
    Errors_Handle "Core_Refresh"
End Sub

' Troca a fonte de dados: pede a nova faixa e recria o gráfico no mesmo lugar, com as mesmas opções.
Public Sub Core_ChangeSource()
    On Error GoTo Fail
    Dim c As Chart, rng As Range
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    On Error Resume Next
    Set rng = Application.InputBox("Selecione a nova tabela (cabeçalho + categorias):", APP_NAME, Type:=8)
    On Error GoTo Fail
    If rng Is Nothing Then Exit Sub
    Core_RebuildFromRange c, rng
    Exit Sub
Fail:
    Errors_Handle "Core_ChangeSource"
End Sub

' Recria o gráfico c a partir da faixa rng, na mesma posição, com as mesmas opções e toggles.
' Anotações (setas) não são preservadas. Retorna True se recriou.
Public Function Core_RebuildFromRange(c As Chart, rng As Range) As Boolean
    On Error GoTo Fail
    Dim co As ChartObject, data As clsChartData, msg As String
    Dim kind As String, l As Double, t As Double, w As Double, h As Double, opts As String, ws As Worksheet, wb As Workbook
    Dim nc As Chart
    Set data = ChartData_FromRange(rng)
    msg = data.Validate()
    If Len(msg) > 0 Then Core_Notify msg, vbExclamation: Exit Function
    Set co = c.Parent
    kind = Registry_ChartType(c)
    l = co.Left: t = co.Top: w = co.Width: h = co.Height
    opts = CStr(Registry_Block(c).Cells(1, 4).Value)
    Set ws = co.Parent: Set wb = ws.Parent
    co.Delete
    CalcSheet_Cleanup wb
    If kind = "waterfall" Then Set co = Waterfall_Insert(ws, data) Else Set co = Stacked_Insert(ws, data, kind)
    co.Left = l: co.Top = t: co.Width = w: co.Height = h
    Set nc = co.Chart
    Registry_Block(nc).Cells(1, 4).Value = opts
    ApplyOption nc, "numfmt", Registry_GetOption(nc, "numfmt", "int")
    ApplyOption nc, "palette", Registry_GetOption(nc, "palette", "padrao")
    If Len(Registry_GetOption(nc, "gap", "")) > 0 Then ApplyOption nc, "gap", Registry_GetOption(nc, "gap", "")
    ApplyToggle nc, "totals", Registry_GetOption(nc, "totals", "0") = "1"
    ApplyToggle nc, "labels", Registry_GetOption(nc, "labels", "1") = "1"
    ApplyToggle nc, "connectors", Registry_GetOption(nc, "connectors", "0") = "1"
    Watchers_Rehook
    co.Select
    Core_SafeRefresh nc
    Core_RebuildFromRange = True
    Exit Function
Fail:
    Errors_Handle "Core_RebuildFromRange"
End Function
