Attribute VB_Name = "modRibbon"
Option Explicit
Private mRibbon As IRibbonUI
Private mHasChart As Boolean        ' cache de Core_HasChart por ciclo de Invalidate
Private mHasChartValid As Boolean
Public Const TUTORIAL_URL As String = "https://israeldias-hub.github.io/ThinQQ/tutorial/tutorial.html"

Public Sub Rx_OnLoad(ribbon As IRibbonUI)
    On Error Resume Next
    Set mRibbon = ribbon
    mHasChartValid = False
End Sub

' Unico ponto de MsgBox do modulo: em build de Teste apenas registra (nunca bloqueia).
Private Sub RxMsg(prompt As String, Optional style As Long = vbInformation)
    If BUILD_CONFIG = "Test" Then
        Loader_Log "RxMsg (teste): " & prompt
    Else
        MsgBox prompt, style, APP_NAME
    End If
End Sub

' Cache de Core_HasChart: uma chamada ao core por ciclo de Invalidate.
Private Function HasChartCached() As Boolean
    If Not mHasChartValid Then
        mHasChart = False
        If Not Loader_CoreWorkbook() Is Nothing Then mHasChart = CBool(Loader_RunCore("Core_HasChart"))
        mHasChartValid = True
    End If
    HasChartCached = mHasChart
End Function

Public Sub Loader_InvalidateRibbon()
    On Error Resume Next
    mHasChartValid = False
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

Private Function CoreOrWarn() As Boolean
    If Loader_CoreWorkbook() Is Nothing Then
        RxMsg "O módulo principal do ThinQQ não está carregado. Use 'Verificar atualizações' ou reinstale.", vbExclamation
        Exit Function
    End If
    CoreOrWarn = True
End Function

Public Sub Rx_OnInsert(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_Insert", control.Tag
End Sub

Public Sub Rx_GetEnabled(control As IRibbonControl, ByRef enabled)
    On Error Resume Next
    enabled = False
    enabled = HasChartCached()
End Sub

Public Sub Rx_OnToggle(control As IRibbonControl, pressed As Boolean)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_SetToggle", control.Tag, pressed
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_GetPressed(control As IRibbonControl, ByRef pressed)
    On Error Resume Next
    pressed = False
    If Not Loader_CoreWorkbook() Is Nothing Then pressed = CBool(Loader_RunCore("Core_GetToggle", control.Tag))
End Sub

Public Sub Rx_OnDropdown(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_SetOption", control.Tag, selectedId
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_GetSelectedId(control As IRibbonControl, ByRef itemId)
    On Error Resume Next
    Dim v As Variant
    v = Loader_RunCore("Core_GetOption", control.Tag)
    If IsEmpty(v) Or Len(CStr(v)) = 0 Then
        Select Case control.Tag
            Case "numfmt": itemId = "int"
            Case "palette": itemId = "padrao"
            Case "gap": itemId = "50"
        End Select
    Else
        itemId = CStr(v)
    End If
End Sub

Public Sub Rx_OnAnnotation(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_AddAnnotation", control.Tag
End Sub

Public Sub Rx_OnRemoveAnnotations(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_RemoveAnnotations"
End Sub

Public Sub Rx_OnRefresh(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_Refresh"
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_OnChangeSource(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_ChangeSource"
End Sub

Public Sub Rx_OnSamples(control As IRibbonControl)
    On Error GoTo Fail
    If Not CoreOrWarn() Then Exit Sub
    Application.Run "'" & Loader_CoreFileName() & "'!Core_InsertSamples"
    Exit Sub
Fail:
    Dim en As Long, ed As String
    en = Err.Number: ed = Err.Description
    Loader_Log "Rx_OnSamples erro " & en & ": " & ed
    If en = 1004 And InStr(1, ed, "Core_InsertSamples", vbTextCompare) > 0 Then
        RxMsg "Os exemplos ainda não estão disponíveis nesta versão. Use 'Verificar atualizações'.", vbInformation
    Else
        RxMsg "Não foi possível inserir os exemplos: " & ed, vbExclamation
    End If
End Sub

' tutorial.html da versão instalada; se não existir, a URL pública.
Public Function Loader_TutorialPath() As String
    Dim ver As String, p As String
    ver = Loader_FindNewestCoreVersion(Loader_VersionsDir())
    If Len(ver) > 0 Then
        p = Loader_VersionsDir() & "\" & ver & "\tutorial.html"
        If Len(Dir(p)) > 0 Then Loader_TutorialPath = p: Exit Function
    End If
    Loader_TutorialPath = TUTORIAL_URL
End Function

Public Sub Rx_OnTutorial(control As IRibbonControl)
    On Error GoTo Fail
    ThisWorkbook.FollowHyperlink Loader_TutorialPath()
    Exit Sub
Fail:
    RxMsg "Não foi possível abrir o tutorial: " & Err.Description, vbExclamation
End Sub

Public Sub Rx_OnCheckUpdates(control As IRibbonControl)
    On Error Resume Next
    Updater_CheckInteractive
End Sub

Public Sub Rx_OnSettings(control As IRibbonControl)
    Dim cur As String, nv As Variant
    On Error GoTo Fail
    cur = Settings_UpdateSource()
    If BUILD_CONFIG = "Test" Then Exit Sub                ' InputBox bloquearia os testes
    nv = Application.InputBox("Fonte de atualização (URL do manifest.json ou caminho de pasta/arquivo):", APP_NAME, cur, Type:=2)
    If VarType(nv) = vbBoolean Then Exit Sub            ' cancelado
    If Len(Trim$(CStr(nv))) = 0 Then
        Settings_Delete KEY_UPDATE_SOURCE
    Else
        Settings_Set KEY_UPDATE_SOURCE, Trim$(CStr(nv))
    End If
    RxMsg "Fonte de atualização: " & Settings_UpdateSource(), vbInformation
    Exit Sub
Fail:
    RxMsg "Não foi possível alterar as configurações: " & Err.Description, vbExclamation
End Sub

Public Sub Rx_OnAbout(control As IRibbonControl)
    On Error GoTo Fail
    RxMsg "ThinQQ" & vbLf & "Loader " & APP_VERSION & "  |  Core " & Loader_CurrentCoreVersion() & vbLf & vbLf & _
           "Gráficos de apresentação para Excel." & vbLf & "Pasta: " & Loader_Root(), vbInformation
    Exit Sub
Fail:
    RxMsg "Não foi possível exibir as informações: " & Err.Description, vbExclamation
End Sub
