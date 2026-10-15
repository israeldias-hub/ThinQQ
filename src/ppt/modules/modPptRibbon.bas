Attribute VB_Name = "modPptRibbon"
Option Explicit
' Faixa do suplemento do PowerPoint (aba ThinQQ, grupo "Graficos do Excel"). Nunca propaga erro.
Private mRibbon As IRibbonUI

Public Sub PptRx_OnLoad(ribbon As IRibbonUI)
    On Error Resume Next
    Set mRibbon = ribbon
End Sub

Public Sub PptRx_OnSync(control As IRibbonControl)
    On Error Resume Next
    If Application.Presentations.Count = 0 Then PptTqq_Msg PPT_MSG_NO_PRES: Exit Sub
    PptTqq_Sync Application.ActivePresentation, True
End Sub

' Itens do dropdown "Ao abrir": modeask, modealways, modenever.
Public Sub PptRx_OnOpenMode(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    On Error Resume Next
    PptTqq_SetOpenMode Mid$(selectedId, 5)
End Sub

Public Sub PptRx_GetOpenMode(control As IRibbonControl, ByRef itemId)
    On Error Resume Next
    itemId = "mode" & PptTqq_OpenMode()
End Sub

Public Sub PptRx_OnAbout(control As IRibbonControl)
    On Error Resume Next
    PptTqq_Msg "ThinQQ para PowerPoint " & APP_VERSION & vbLf & "Atualiza dados e setas dos gráficos do ThinQQ copiados como gráfico vinculado ao abrir a apresentação (aba ThinQQ > Ao abrir).", vbInformation
End Sub
