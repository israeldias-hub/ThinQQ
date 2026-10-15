Attribute VB_Name = "modPptAuto"
Option Explicit
' Atualizacao automatica dos elementos do ThinQQ no PowerPoint (spec 2026-10-15 secao 5.1). Opcoes em
' HKCU\Software\VB and VBA Program Settings\ThinQQ\PowerPoint: AutoUpdate (padrao "1": o motor age nas formas marcadas TQQ_AUTO) e AutoNew
' (padrao "0": o vinculo nasce manual, como no think-cell). No build Test so memoria (nunca o registro).
Private Const PPTAUTO_SECTION As String = "PowerPoint"
Private mTestEnabled As Variant                 ' gancho: Empty = padrao
Private mTestNewAuto As Variant

Public Sub PptAuto_SetFlagsForTest(enabled As Variant, newAuto As Variant)
    If BUILD_CONFIG <> "Test" Then Exit Sub
    mTestEnabled = enabled
    mTestNewAuto = newAuto
End Sub

Public Sub PptAuto_ResetTestHooks()
    mTestEnabled = Empty
    mTestNewAuto = Empty
End Sub

Public Function PptAuto_Enabled() As Boolean
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then
        If IsEmpty(mTestEnabled) Then PptAuto_Enabled = True Else PptAuto_Enabled = CBool(mTestEnabled)
        Exit Function
    End If
    PptAuto_Enabled = (GetSetting(APP_NAME, PPTAUTO_SECTION, "AutoUpdate", "1") <> "0")
End Function

Public Sub PptAuto_SetEnabled(v As Boolean)
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then mTestEnabled = v: Exit Sub
    SaveSetting APP_NAME, PPTAUTO_SECTION, "AutoUpdate", IIf(v, "1", "0")
End Sub

Public Function PptAuto_NewLinksAuto() As Boolean
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then
        If IsEmpty(mTestNewAuto) Then PptAuto_NewLinksAuto = False Else PptAuto_NewLinksAuto = CBool(mTestNewAuto)
        Exit Function
    End If
    PptAuto_NewLinksAuto = (GetSetting(APP_NAME, PPTAUTO_SECTION, "AutoNew", "0") = "1")
End Function

Public Sub PptAuto_SetNewLinksAuto(v As Boolean)
    On Error Resume Next
    If BUILD_CONFIG = "Test" Then mTestNewAuto = v: Exit Sub
    SaveSetting APP_NAME, PPTAUTO_SECTION, "AutoNew", IIf(v, "1", "0")
End Sub
