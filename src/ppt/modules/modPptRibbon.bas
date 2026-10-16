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
    PptTqq_Msg "ThinQQ para PowerPoint " & APP_VERSION & vbLf & "Cria gráficos do ThinQQ direto no slide (Inserir gráfico; os dados ficam numa pasta do Excel ao lado da apresentação) e atualiza dados e setas dos gráficos do ThinQQ vinculados: ao abrir a apresentação (aba ThinQQ > Ao abrir), pelo botão ou automaticamente (Gráfico selecionado > Atualizar automaticamente).", vbInformation
End Sub

' ---- Grupo "Grafico selecionado" (1.6.0): estado segue a selecao (clsPptEvents.WindowSelectionChange invalida). ----

Public Sub PptRx_InvalidateElem()
    Dim id As Variant
    On Error Resume Next
    If mRibbon Is Nothing Then Exit Sub
    mRibbon.InvalidateControl "tglPptAuto"
    mRibbon.InvalidateControl "btnPptUpdateSel"
    mRibbon.InvalidateControl "btnPptEditData"
    ' grupo "Grafico do ThinQQ" (1.9.0): estado e capacidades vem da etiqueta TQQ_OPTS da forma selecionada
    For Each id In Array("tglPptLabels", "tglPptTotals", "tglPptConnectors", "ddPptNumFmt", "ddPptLegend", "btnPptCagr", "btnPptDiff", "btnPptClearAnn")
        mRibbon.InvalidateControl CStr(id)
    Next
End Sub

' Ligado so com pelo menos um grafico do ThinQQ (de um slide) selecionado.
Public Sub PptRx_GetElemEnabled(control As IRibbonControl, ByRef enabled)
    On Error Resume Next
    enabled = False
    enabled = (PptTqq_SelectedTagged().Count > 0)
End Sub

' Pressionado so se TODOS os graficos do ThinQQ selecionados sao automaticos (selecao mista: solto; clicar liga todos).
Public Sub PptRx_GetAuto(control As IRibbonControl, ByRef pressed)
    On Error Resume Next
    pressed = False
    pressed = PptTqq_ShapesAuto(PptTqq_SelectedTagged())
End Sub

' Ligar ja atualiza uma vez (como o "Update Automatically" do think-cell).
Public Sub PptRx_OnAuto(control As IRibbonControl, pressed As Boolean)
    On Error Resume Next
    PptTqq_ToggleAuto PptTqq_SelectedTagged(), pressed
    PptRx_InvalidateElem
End Sub

Public Sub PptRx_OnUpdateSel(control As IRibbonControl)
    On Error Resume Next
    PptTqq_UpdateSelected PptTqq_SelectedTagged()
End Sub

' ---- Grupo "Inserir grafico" e "Editar dados" (1.8.0): o tipo vem do tag do botao (os mesmos 21 do Excel). ----

Public Sub PptRx_OnInsertChart(control As IRibbonControl)
    Dim pres As Object
    On Error Resume Next
    If Application.Presentations.Count = 0 Then PptTqq_Msg PPT_MSG_NO_PRES: Exit Sub
    Set pres = Application.ActivePresentation
    PptTqq_InsertChartIn pres, PptTqq_ActiveSlide(pres), control.Tag
End Sub

' A apresentacao e a da forma selecionada (sem selecao: o aviso de selecionar).
Public Sub PptRx_OnEditData(control As IRibbonControl)
    Dim shapes As Collection
    On Error Resume Next
    Set shapes = PptTqq_SelectedTagged()
    If shapes.Count = 0 Then PptTqq_Msg PPT_MSG_SELECT: Exit Sub
    PptTqq_EditShapes shapes(1).Parent.Parent, shapes
End Sub

' ---- Grupo "Grafico do ThinQQ" (1.9.0): a opcao vem do tag do controle; a apresentacao, da forma selecionada (PptTqq_CommandShapes). ----

' Tag do controle ("" sem controle). Nunca propaga.
Private Function PptRx_TagOf(control As IRibbonControl) As String
    On Error Resume Next
    PptRx_TagOf = CStr(control.Tag)
    Err.Clear
End Function

Public Sub PptRx_GetOpt(control As IRibbonControl, ByRef pressed)
    On Error Resume Next
    pressed = False
    pressed = (PptTqq_OptValue(PptTqq_SelectedOpts(), PptRx_TagOf(control)) = "1")
End Sub

Public Sub PptRx_GetOptEnabled(control As IRibbonControl, ByRef enabled)
    On Error Resume Next
    enabled = False
    enabled = PptTqq_OptCap(PptTqq_SelectedOpts(), PptRx_TagOf(control))
End Sub

Public Sub PptRx_OnToggleOpt(control As IRibbonControl, pressed As Boolean)
    On Error Resume Next
    PptTqq_CommandShapes Nothing, PptTqq_SelectedTagged(), "toggle", PptRx_TagOf(control)
End Sub

' Itens dos dropdowns: id = tag + valor ("numfmtpct0", "legendright").
Public Sub PptRx_GetOptItem(control As IRibbonControl, ByRef itemId)
    Dim t As String
    On Error Resume Next
    t = PptRx_TagOf(control)
    itemId = t & PptTqq_OptValue(PptTqq_SelectedOpts(), t)
End Sub

Public Sub PptRx_OnOptItem(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    Dim t As String
    On Error Resume Next
    t = PptRx_TagOf(control)
    PptTqq_CommandShapes Nothing, PptTqq_SelectedTagged(), t, Mid$(selectedId, Len(t) + 1)
End Sub

Public Sub PptRx_OnArrow(control As IRibbonControl)
    On Error Resume Next
    PptTqq_CommandShapes Nothing, PptTqq_SelectedTagged(), PptRx_TagOf(control), ""
End Sub

Public Sub PptRx_OnClearAnn(control As IRibbonControl)
    On Error Resume Next
    PptTqq_CommandShapes Nothing, PptTqq_SelectedTagged(), "clear", ""
End Sub
