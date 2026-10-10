Attribute VB_Name = "modRibbon"
Option Explicit
Private mRibbon As IRibbonUI
Public Const TUTORIAL_URL As String = "https://github.com/israeldias-hub/ThinQQ/blob/main/docs/tutorial/tutorial.html"

Public Sub Rx_OnLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
End Sub

Public Sub Loader_InvalidateRibbon()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

Private Function CoreOrWarn() As Boolean
    If Loader_CoreWorkbook() Is Nothing Then
        MsgBox "O módulo principal do ThinQQ não está carregado. Use 'Verificar atualizações' ou reinstale.", vbExclamation, APP_NAME
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
    If Not Loader_CoreWorkbook() Is Nothing Then enabled = CBool(Loader_RunCore("Core_HasChart"))
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
    If Not CoreOrWarn() Then Exit Sub
    On Error GoTo Fail
    Application.Run "'" & Loader_CoreFileName() & "'!Core_InsertSamples"
    Exit Sub
Fail:
    ' O core instalado pode ainda não ter os exemplos: aviso amigável, sem erro modal.
    Loader_Log "Rx_OnSamples erro " & Err.Number & ": " & Err.Description
    MsgBox "Os exemplos ainda não estão disponíveis nesta versão. Use 'Verificar atualizações'.", vbInformation, APP_NAME
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
    MsgBox "Não foi possível abrir o tutorial: " & Err.Description, vbExclamation, APP_NAME
End Sub

Public Sub Rx_OnCheckUpdates(control As IRibbonControl)
    On Error Resume Next
    Updater_CheckInteractive
End Sub

Public Sub Rx_OnSettings(control As IRibbonControl)
    Dim cur As String, nv As Variant
    cur = Settings_UpdateSource()
    nv = Application.InputBox("Fonte de atualização (URL do manifest.json ou caminho de pasta/arquivo):", APP_NAME, cur, Type:=2)
    If VarType(nv) = vbBoolean Then Exit Sub            ' cancelado
    If Len(Trim$(CStr(nv))) = 0 Then
        Settings_Delete KEY_UPDATE_SOURCE
    Else
        Settings_Set KEY_UPDATE_SOURCE, Trim$(CStr(nv))
    End If
    MsgBox "Fonte de atualização: " & Settings_UpdateSource(), vbInformation, APP_NAME
End Sub

Public Sub Rx_OnAbout(control As IRibbonControl)
    MsgBox "ThinQQ" & vbLf & "Loader " & APP_VERSION & "  |  Core " & Loader_CurrentCoreVersion() & vbLf & vbLf & _
           "Gráficos de apresentação para Excel." & vbLf & "Pasta: " & Loader_Root(), vbInformation, APP_NAME
End Sub
