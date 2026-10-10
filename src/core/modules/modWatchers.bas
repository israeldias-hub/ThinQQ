Attribute VB_Name = "modWatchers"
Option Explicit
Private mWatchers As Collection
Private mAppEvents As clsAppEvents
Private mLastHasChart As Boolean
Private mLoaderName As String

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
    Dim ws As Object, co As ChartObject, w As clsChartWatcher
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
            On Error GoTo 0
        End If
    Next
End Sub

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
