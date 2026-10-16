Attribute VB_Name = "modWatchers"
Option Explicit
Private mWatchers As Collection
Private mAppEvents As clsAppEvents
Private mLastHasChart As Boolean
Private mLoaderName As String
Private mRehooking As Boolean       ' reentrada: Core_SafeRefresh pode disparar eventos que chamam Rehook

Public Sub Watchers_Start()
    On Error Resume Next
    Set mAppEvents = New clsAppEvents
    Set mAppEvents.App = Application
    Watchers_Rehook
End Sub

Public Sub Watchers_Stop()
    Set mAppEvents = Nothing
    Set mWatchers = Nothing
    mLastHasChart = False
    mLoaderName = ""
End Sub

Public Sub Watchers_SetLoader(name As String)
    mLoaderName = name
End Sub

' Reassocia observadores a todos os gráficos ThinQQ da planilha ativa. Gráficos de planilhas inativas
' não têm observador: são atualizados quando a planilha é ativada (SheetActivate -> Rehook -> eventos).
Public Sub Watchers_Rehook()
    Dim ws As Object, co As ChartObject, w As clsChartWatcher, wbk As Workbook
    Set mWatchers = New Collection
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If TypeName(ws) <> "Worksheet" Then Exit Sub
    For Each co In ws.ChartObjects
        If Left$(co.Name, Len(ID_PREFIX)) = ID_PREFIX Then
            On Error Resume Next
            Set w = New clsChartWatcher
            Set w.Cht = co.Chart
            If Err.Number = 0 Then mWatchers.Add w
            Err.Clear
            Core_NotifyCopiedChartOnce co           ' planilha copiada: o grafico divide id e bloco com o original (aviso unico)
            Err.Clear
            On Error GoTo 0
            ' redesenha anotacoes/rotulos do grafico que acabou de ser observado (planilha recem-ativada)
            If Not mRehooking Then
                mRehooking = True
                Core_SafeRefresh co.Chart
                mRehooking = False
            End If
        End If
    Next
    ' molduras de intervalos vinculados (1.7.0) cujo grafico nao existe mais: somem aqui
    On Error Resume Next
    Set wbk = ws.Parent
    If Not wbk Is Nothing Then LinkFrame_CleanOrphans wbk
    Err.Clear
End Sub

' Estado dos observadores (para quem precisa restaura-lo, como os testes da faixa): ligados (eventos do aplicativo) e o loader registrado.
Public Function Watchers_IsActive() As Boolean
    Watchers_IsActive = Not mAppEvents Is Nothing
End Function

Public Function Watchers_LoaderName() As String
    Watchers_LoaderName = mLoaderName
End Function

' True se o grafico (pelo nome do ChartObject, da planilha e da pasta) tem observador. Nunca propaga erro.
Public Function Watchers_IsWatched(cht As Chart) As Boolean
    Dim w As clsChartWatcher, key As String
    On Error Resume Next
    If mWatchers Is Nothing Then Exit Function
    key = cht.Parent.Parent.Parent.Name & "|" & cht.Parent.Parent.Name & "|" & cht.Parent.Name
    If Err.Number <> 0 Then Exit Function
    For Each w In mWatchers
        If w.Cht.Parent.Parent.Parent.Name & "|" & w.Cht.Parent.Parent.Name & "|" & w.Cht.Parent.Name = key Then
            If Err.Number = 0 Then Watchers_IsWatched = True: Exit Function
        End If
        Err.Clear
    Next
End Function

Public Function Watchers_Count() As Long
    If mWatchers Is Nothing Then Exit Function
    Watchers_Count = mWatchers.Count
End Function

' Invalida a faixa quando entra/sai de um gráfico ThinQQ (ou troca de gráfico).
Public Sub Watchers_NotifySelection()
    Dim hasChart As Boolean
    On Error Resume Next
    hasChart = Core_HasChart()
    If hasChart Or hasChart <> mLastHasChart Then
        mLastHasChart = hasChart
        If Len(mLoaderName) > 0 Then
            Application.Run "'" & mLoaderName & "'!Loader_InvalidateRibbon"
        End If
    End If
End Sub
