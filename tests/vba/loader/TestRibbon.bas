Attribute VB_Name = "TestRibbon"
Option Explicit

Private Function LoadRibbon() As Object
    Dim d As Object
    Set d = CreateObject("MSXML2.DOMDocument.6.0")
    d.async = False
    d.setProperty "SelectionLanguage", "XPath"
    d.setProperty "SelectionNamespaces", "xmlns:u='http://schemas.microsoft.com/office/2009/07/customui'"
    d.Load RepoPath() & "\src\loader\ribbon\customUI14.xml"
    AssertEqual 0, d.parseError.errorCode, "XML malformado: " & d.parseError.reason
    Set LoadRibbon = d
End Function

' Valor do atributo, ou "" se o no e Nothing ou nao tem o atributo (nunca erro 91: mensagens de falha nao podem quebrar o teste).
Private Function AttrText(n As Object, attr As String) As String
    Dim a As Object
    If n Is Nothing Then Exit Function
    Set a = n.Attributes.getNamedItem(attr)
    If Not a Is Nothing Then AttrText = a.Text
End Function

' id do controle para mensagens de falha; "<elemento sem id>" quando nao ha id.
Private Function NodeId(n As Object) As String
    NodeId = AttrText(n, "id")
    If Len(NodeId) = 0 Then
        If n Is Nothing Then NodeId = "<nada>" Else NodeId = "<" & n.baseName & " sem id>"
    End If
End Function

Private Function RibbonModuleCode() As String
    Dim cm As Object
    Set cm = ThisWorkbook.VBProject.VBComponents("modRibbon").CodeModule
    RibbonModuleCode = cm.Lines(1, cm.CountOfLines)
End Function

' Tipos dos parametros do Public Sub/Function "name", normalizados: "[byref ]tipo" separados por ",".
' Retorna "?" se o procedimento nao existe como Public.
Private Function SigOf(code As String, name As String) As String
    Dim lines() As String, i As Long, ln As String, p As Long, q As Long
    Dim ps() As String, j As Long, one As String, byRefFlag As Boolean, t As String, out As String
    lines = Split(Replace(code, vbCr, ""), vbLf)
    SigOf = "?"
    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If Left$(ln, 11) = "Public Sub " Or Left$(ln, 16) = "Public Function " Then
            p = InStr(ln, " " & name & "(")
            If p > 0 Then
                q = InStrRev(ln, ")")
                one = Mid$(ln, p + Len(name) + 2, q - p - Len(name) - 2)
                out = ""
                If Len(Trim$(one)) > 0 Then
                    ps = Split(one, ",")
                    For j = 0 To UBound(ps)
                        t = LCase$(Trim$(ps(j)))
                        byRefFlag = (InStr(t, "byref ") > 0)
                        t = Replace(Replace(t, "byref ", ""), "byval ", "")
                        If InStr(t, " as ") > 0 Then
                            t = Trim$(Mid$(t, InStr(t, " as ") + 4))
                        Else
                            t = "variant"
                        End If
                        If j > 0 Then out = out & ","
                        out = out & IIf(byRefFlag, "byref ", "") & t
                    Next
                End If
                SigOf = out
                Exit Function
            End If
        End If
    Next
End Function

Private Function ExpectedSig(elem As String, attr As String) As String
    Select Case attr
        Case "onLoad": ExpectedSig = "iribbonui"
        Case "onAction"
            Select Case elem
                Case "button": ExpectedSig = "iribboncontrol"
                Case "toggleButton", "checkBox": ExpectedSig = "iribboncontrol,boolean"
                Case "dropDown": ExpectedSig = "iribboncontrol,string,integer"
            End Select
        Case "getEnabled", "getPressed", "getSelectedItemID": ExpectedSig = "iribboncontrol,byref *"
        Case "getItemCount": ExpectedSig = "iribboncontrol,byref *"
        Case "getItemLabel", "getItemID": ExpectedSig = "iribboncontrol,integer,byref *"
    End Select
End Function

Private Function SigMatches(expected As String, actual As String) As Boolean
    Dim e() As String, a() As String, i As Long
    e = Split(expected, ","): a = Split(actual, ",")
    If UBound(e) <> UBound(a) Then Exit Function
    For i = 0 To UBound(e)
        If e(i) = "byref *" Then
            If Left$(a(i), 6) <> "byref " Then Exit Function
        ElseIf e(i) <> a(i) Then
            Exit Function
        End If
    Next
    SigMatches = True
End Function

Public Sub Test_Ribbon_Xml_WellFormedAndNamespace()
    Dim d As Object
    Set d = LoadRibbon()
    AssertEqual "http://schemas.microsoft.com/office/2009/07/customui", d.documentElement.namespaceURI
    AssertEqual "customUI", d.documentElement.baseName
End Sub

Public Sub Test_Ribbon_Xml_HasAllCallbacks()
    Dim xml As String, names As Variant, i As Long
    xml = Transport_ReadTextFile(RepoPath() & "\src\loader\ribbon\customUI14.xml")
    names = Array("Rx_OnLoad", "Rx_OnInsert", "Rx_GetEnabled", "Rx_OnToggle", "Rx_GetPressed", "Rx_OnDropdown", _
                  "Rx_GetSelectedId", "Rx_OnAnnotation", "Rx_OnRemoveAnnotations", "Rx_OnRefresh", "Rx_OnChangeSource", _
                  "Rx_OnSamples", "Rx_OnTutorial", "Rx_OnCheckUpdates", "Rx_OnSettings", "Rx_OnAbout", "Rx_OnSameScale", "Rx_OnResetScale", _
                  "Rx_OnAxisBreak", "Rx_OnAxisBreakOff", "Rx_OnRoundRange", "Rx_OnPptImage", "Rx_OnPptChart", "Rx_OnPptSync", "Rx_OnDelta", "Rx_OnDeltaFormat", _
                  "Rx_GetPptInsert", "Rx_OnPptInsert", "Rx_OnPptFrames", "Rx_OnPptLinkText", "Rx_OnLinkTitle", _
                  "Rx_StyleCount", "Rx_StyleLabel", "Rx_StyleId", "Rx_GetStyleSel", "Rx_OnStyle", "Rx_OnStyles")
    For i = LBound(names) To UBound(names)
        AssertTrue InStr(xml, names(i)) > 0, "callback ausente no XML: " & names(i)
    Next
    AssertTrue InStr(xml, "tag=""butterfly""") > 0
    AssertTrue InStr(xml, "tag=""waterfall""") > 0
    AssertTrue InStr(xml, "tag=""gantt""") > 0
    AssertTrue InStr(xml, "tag=""numfmt""") > 0
End Sub

' Todo callback do XML existe como Public em modRibbon com a assinatura certa para o tipo de controle.
Public Sub Test_Ribbon_Callbacks_ExistWithRightSignature()
    Dim d As Object, code As String, attrs As Variant, i As Long, n As Object, a As Object
    Dim cb As String, sig As String, exp As String, count As Long
    Set d = LoadRibbon()
    code = RibbonModuleCode()
    attrs = Array("onLoad", "onAction", "getEnabled", "getPressed", "getSelectedItemID", "getItemCount", "getItemLabel", "getItemID")
    For Each n In d.SelectNodes("//*[@onLoad or @onAction or @getEnabled or @getPressed or @getSelectedItemID or @getItemCount or @getItemLabel or @getItemID]")
        For i = LBound(attrs) To UBound(attrs)
            Set a = n.Attributes.getNamedItem(CStr(attrs(i)))
            If Not a Is Nothing Then
                cb = a.Text
                sig = SigOf(code, cb)
                AssertTrue sig <> "?", "callback sem Public Sub em modRibbon: " & cb
                exp = ExpectedSig(n.baseName, CStr(attrs(i)))
                AssertTrue Len(exp) > 0, "combinacao inesperada " & n.baseName & "/" & attrs(i)
                AssertTrue SigMatches(exp, sig), "assinatura de " & cb & " em " & n.baseName & "/" & attrs(i) & ": esperado [" & exp & "] obtido [" & sig & "]"
                count = count + 1
            End If
        Next
    Next
    AssertTrue count >= 30, "poucos callbacks verificados: " & count
End Sub

' Ids unicos E validos: o id de todo controle e de todo item segue o padrao de ID de XML (comeca por letra ou "_"; so letras, digitos e "_").
' Um id invalido (ex.: o item "30" do dropdown Espaco, na Fase 1) faz o Excel descartar a faixa ThinQQ INTEIRA, em silencio (medido abrindo o
' customUI14 num xlsx minimo: com os ids "30", "50", "80" e "120" a aba nao aparece; com "gap30" etc. aparece).
Public Sub Test_Ribbon_ControlIds_Unique()
    Dim d As Object, n As Object, seen As Object, id As String, re As Object
    Set d = LoadRibbon()
    Set seen = CreateObject("Scripting.Dictionary")
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Za-z_][A-Za-z0-9_]*$"
    For Each n In d.SelectNodes("//*[@id]")
        id = n.Attributes.getNamedItem("id").Text
        AssertFalse seen.Exists(id), "id duplicado: " & id
        AssertTrue re.Test(id), "id invalido (o Excel descarta a aba inteira, em silencio): '" & id & "'"
        seen.Add id, 1
    Next
    AssertTrue seen.Count > 20
End Sub

Private Function GetStyleSelForTest() As String
    Dim sel As Variant
    Loader_InvalidateRibbon
    Rx_GetStyleSel Nothing, sel
    GetStyleSelForTest = CStr(sel)
End Function

Private Function ItemIds(dd As Object) As String
    Dim it As Object, s As String
    For Each it In dd.SelectNodes("u:item")
        s = s & IIf(Len(s) > 0, ",", "") & it.Attributes.getNamedItem("id").Text
    Next
    ItemIds = s
End Function

Public Sub Test_Ribbon_Dropdowns_ItemsMatchCoreOptions()
    Dim d As Object, dd As Object, it As Object, v As String
    Set d = LoadRibbon()
    AssertEqual "int,dec1,dec2,pct0,pct1,mil,mi", ItemIds(d.SelectSingleNode("//u:dropDown[@tag='numfmt']"))
    AssertTrue d.SelectSingleNode("//u:dropDown[@id='ddPalette']") Is Nothing, "Paleta virou Estilo (itens vindos do core)"
    Set dd = d.SelectSingleNode("//u:dropDown[@tag='gap']")
    AssertTrue Not dd Is Nothing
    For Each it In dd.SelectNodes("u:item")
        v = it.Attributes.getNamedItem("id").Text
        AssertEqual "gap", Left$(v, 3), "o id do item de gap e 'gap' + numero (id de XML nao comeca por digito): " & v
        v = Rx_OptionValueFor("gap", v)
        AssertTrue IsNumeric(v), "item de gap nao numerico: " & v
        AssertTrue CLng(v) >= 0 And CLng(v) <= 500, "gap fora de 0..500: " & v
    Next
    AssertEqual "30,50,80,120", Replace(ItemIds(dd), "gap", "")
    AssertEqual 4, d.SelectNodes("//u:dropDown").Length
End Sub

' Toda tag da faixa existe no core REAL (loader e core de teste abertos juntos): tipo de grafico = Kinds_IsValid, toggle e dropdown =
' recurso conhecido da matriz (Kinds_FeatureLabel), anotacao = recurso suportado por algum tipo (Kinds_Supports). As contagens e as listas
' fixas ficam em Test_Ribbon_Menus_Structure e Test_Ribbon_Toggles_And_Dropdowns_Phase2.
Public Sub Test_Ribbon_Tags_MatchCoreKeys()
    Dim d As Object, n As Object, t As String, k As Object, found As Boolean
    Set d = LoadRibbon()
    AssertEqual 21, d.SelectNodes("//u:button[@onAction='Rx_OnInsert']").Length
    For Each n In d.SelectNodes("//u:button[@onAction='Rx_OnInsert']")
        t = n.Attributes.getNamedItem("tag").Text
        AssertTrue CBool(Loader_RunCore("Kinds_IsValid", t)), "tipo desconhecido do core: " & t
    Next
    For Each n In d.SelectNodes("//u:toggleButton | //u:dropDown")
        t = n.Attributes.getNamedItem("tag").Text
        AssertTrue Len(CStr(Loader_RunCore("Kinds_FeatureLabel", t))) > 0, "recurso desconhecido do core: " & t
    Next
    For Each n In d.SelectNodes("//u:button[@onAction='Rx_OnAnnotation']")
        t = n.Attributes.getNamedItem("tag").Text
        found = False
        For Each k In d.SelectNodes("//u:button[@onAction='Rx_OnInsert']")
            If CBool(Loader_RunCore("Kinds_Supports", k.Attributes.getNamedItem("tag").Text, t)) Then found = True: Exit For
        Next
        AssertTrue found, "nenhum tipo suporta a anotacao " & t
    Next
    ' Botoes com tag de recurso (Corte de eixo): a tag e um recurso da matriz (Core_Supports decide o cinza); "breakoff" (Remover corte de
    ' eixo) e uma tag de ESTADO que o core conhece (ativo so com corte; aqui, sem grafico, cinza)
    For Each n In d.SelectNodes("//u:button[@tag][not(@onAction='Rx_OnInsert')][not(@onAction='Rx_OnAnnotation')]")
        t = n.Attributes.getNamedItem("tag").Text
        If t = "breakoff" Then
            AssertFalse CBool(Loader_RunCore("Core_Supports", t)), "Remover corte de eixo sem grafico"
        Else
            AssertTrue Len(CStr(Loader_RunCore("Kinds_FeatureLabel", t))) > 0, "recurso desconhecido do core: " & t
        End If
        AssertEqual "Rx_GetEnabled", AttrText(n, "getEnabled"), "botao com tag de recurso sem getEnabled: " & NodeId(n)
    Next
End Sub

' Menus (customUI14): 5 menus de insercao com os 21 botoes + "Mais opcoes", "Setas" e "PowerPoint". Um atributo invalido em filho de menu
' faz o Excel descartar a aba inteira, sem aviso: nenhum filho de menu leva 'size'.
Public Sub Test_Ribbon_Menus_Structure()
    Dim d As Object, m As Object, kids As Object, n As Object, t As String
    Set d = LoadRibbon()
    AssertEqual 8, d.SelectNodes("//u:menu").Length, "5 menus de insercao + Mais opcoes + Setas + PowerPoint"
    AssertEqual 5, d.SelectNodes("//u:group[@id='grpInsert']/u:menu[@size='large']").Length
    AssertEqual 0, d.SelectNodes("//u:menu/*[@size]").Length, "filho de menu com 'size' (invalido: o Excel descarta a aba inteira sem aviso)"
    For Each m In d.SelectNodes("//u:group[@id='grpInsert']/u:menu")
        Set kids = m.SelectNodes("u:button")
        AssertTrue kids.Length >= 2, "menu " & m.Attributes.getNamedItem("id").Text & " com poucos itens"
        For Each n In kids
            AssertEqual "Rx_OnInsert", n.Attributes.getNamedItem("onAction").Text
            t = n.Attributes.getNamedItem("tag").Text
            AssertTrue InStr(",stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo,mekko,mekkounits,gantt,pie,doughnut,butterfly,scatter,bubble,waterfallpair,", "," & t & ",") > 0, "kind invalido: " & t
        Next
    Next
    AssertEqual 21, d.SelectNodes("//u:button[@onAction='Rx_OnInsert']").Length
    AssertTrue d.SelectSingleNode("//u:menu[@id='mnuMore']/@getEnabled") Is Nothing, "Mais opcoes contem botoes sempre ativos"
    AssertEqual 7, d.SelectNodes("//u:menu[@id='mnuMore']/u:toggleButton").Length
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnSameScale'][@onAction='Rx_OnSameScale']") Is Nothing
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnResetScale'][@onAction='Rx_OnResetScale']") Is Nothing
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuArrows']/u:button[@tag='bracket']") Is Nothing
    AssertEqual "Rx_GetEnabled", d.SelectSingleNode("//u:menu[@id='mnuArrows']").Attributes.getNamedItem("getEnabled").Text
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnBreak'][@tag='break'][@onAction='Rx_OnAxisBreak'][@getEnabled='Rx_GetEnabled']") Is Nothing
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnBreakOff'][@tag='breakoff'][@onAction='Rx_OnAxisBreakOff']") Is Nothing
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnRoundRange'][@onAction='Rx_OnRoundRange']") Is Nothing
    AssertTrue d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnRoundRange']/@getEnabled") Is Nothing, "Arredondar intervalo age sobre celulas: sempre ativo"
    AssertTrue d.SelectSingleNode("//u:menu[@id='mnuPpt']/@getEnabled") Is Nothing, "PowerPoint sempre ativo: Atualizar graficos no PowerPoint nao exige grafico"
    AssertTrue d.SelectSingleNode("//u:menu[@id='mnuPpt']/@tag") Is Nothing
    AssertEqual 6, d.SelectNodes("//u:menu[@id='mnuPpt']/u:button").Length
    AssertEqual "Rx_GetEnabled", AttrText(d.SelectSingleNode("//u:button[@id='btnPptImage']"), "getEnabled"), "copiar exige grafico"
    AssertEqual "Rx_GetEnabled", AttrText(d.SelectSingleNode("//u:button[@id='btnPptChart']"), "getEnabled")
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:button[@id='btnPptSync'][@onAction='Rx_OnPptSync']") Is Nothing
    AssertTrue d.SelectSingleNode("//u:button[@id='btnPptSync']/@getEnabled") Is Nothing, "Atualizar graficos no PowerPoint: sempre ativo"
    AssertEqual "btnPptLinkText", AttrText(d.SelectSingleNode("//u:button[@id='btnPptChart']/following-sibling::*[1]"), "id"), "1.9.0: vincular texto logo depois de copiar"
    AssertEqual "btnPptTable", AttrText(d.SelectSingleNode("//u:button[@id='btnPptLinkText']/following-sibling::*[1]"), "id"), "1.10.0: tabela vinculada logo depois do texto"
    AssertEqual "msepPpt", AttrText(d.SelectSingleNode("//u:button[@id='btnPptTable']/following-sibling::*[1]"), "id")
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:button[@id='btnPptImage'][@onAction='Rx_OnPptImage']") Is Nothing
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:button[@id='btnPptChart'][@onAction='Rx_OnPptChart']") Is Nothing
    ' ordem do grupo: ... Mais opcoes, Setas, PowerPoint, separador, Atualizar
    AssertEqual "mnuPpt", AttrText(d.SelectSingleNode("//u:menu[@id='mnuArrows']/following-sibling::*[1]"), "id"), "PowerPoint logo depois de Setas"
    ' ordem dos itens novos de Mais tipos: depois da Rosca
    AssertEqual "btnDoughnut,btnButterfly,btnScatter,btnBubble", AttrText(d.SelectSingleNode("//u:button[@id='btnDoughnut']"), "id") & "," & _
        AttrText(d.SelectSingleNode("//u:button[@id='btnDoughnut']/following-sibling::*[1]"), "id") & "," & _
        AttrText(d.SelectSingleNode("//u:button[@id='btnDoughnut']/following-sibling::*[2]"), "id") & "," & _
        AttrText(d.SelectSingleNode("//u:button[@id='btnDoughnut']/following-sibling::*[3]"), "id")
    ' Cascata comparativa: terceiro item do menu Cascata; as setas de variacao (2 alternadores) no menu Setas, antes do separador
    AssertEqual "btnWaterfall,btnBarWaterfall,btnWaterfallPair", AttrText(d.SelectSingleNode("//u:menu[@id='mnuWaterfall']/u:button[1]"), "id") & "," & AttrText(d.SelectSingleNode("//u:menu[@id='mnuWaterfall']/u:button[2]"), "id") & "," & AttrText(d.SelectSingleNode("//u:menu[@id='mnuWaterfall']/u:button[3]"), "id")
    AssertEqual 2, d.SelectNodes("//u:menu[@id='mnuArrows']/u:toggleButton").Length
    AssertEqual "msepArrows", AttrText(d.SelectSingleNode("//u:toggleButton[@id='tglPairAbs']/following-sibling::*[1]"), "id"), "as setas de variacao antes do separador"
End Sub

' Rotulos, tags, icones e textos de ajuda exatos dos controles da cascata comparativa (o tutorial e o Test-Docs dependem dos rotulos).
Public Sub Test_Ribbon_PairWaterfall_LabelsAndTips()
    Dim d As Object, spec As Variant, i As Long, p() As String, n As Object
    Set d = LoadRibbon()
    spec = Array("btnWaterfallPair|Cascata comparativa|waterfallpair|tqq_waterfallpair16", "tglPairDiff|Setas de variação|pairdiff|tqq_arrows16", "tglPairAbs|Variação em valor|pairabs|tqq_waterfallpair16")
    For i = LBound(spec) To UBound(spec)
        p = Split(spec(i), "|")
        Set n = d.SelectSingleNode("//*[@id='" & p(0) & "']")
        AssertTrue Not n Is Nothing, "controle ausente: " & p(0)
        AssertEqual p(1), AttrText(n, "label"), "rotulo de " & p(0)
        AssertEqual p(2), AttrText(n, "tag"), "tag de " & p(0)
        AssertEqual p(3), AttrText(n, "image"), "icone de " & p(0)
        AssertTrue Len(AttrText(n, "supertip")) > 20, "sem supertip: " & p(0)
    Next
    AssertEqual "Nome dentro das barras", AttrText(d.SelectSingleNode("//u:dropDown[@tag='legend']/u:item[@id='inbar']"), "label")
    AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='tglTotals']"), "supertip"), "cascata comparativa") > 0, "Totais explica o percentual da comparativa"
    AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='tglPairDiff']"), "supertip"), "n/d") > 0
    AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='btnWaterfallPair']"), "supertip"), "exatamente 2 colunas de valores") > 0
End Sub

' Delta entre barras (Parte B): rotulos, tags, posicao no menu Setas e supertips exatos (o tutorial, a matriz e o Test-Docs dependem dos rotulos).
Public Sub Test_Ribbon_Delta_LabelsAndTips()
    Dim d As Object, n As Object, m As Object
    Set d = LoadRibbon()
    Set n = d.SelectSingleNode("//u:menu[@id='mnuArrows']/u:button[@id='btnDelta']")
    AssertTrue Not n Is Nothing, "btnDelta no menu Setas"
    AssertEqual "Delta clicando nas barras", AttrText(n, "label"): AssertEqual "delta", AttrText(n, "tag")
    AssertEqual "Rx_OnDelta", AttrText(n, "onAction"): AssertEqual "Rx_GetEnabled", AttrText(n, "getEnabled")
    AssertEqual "btnDelta", AttrText(d.SelectSingleNode("//u:button[@id='btnDiff']/following-sibling::*[1]"), "id"), "logo depois da Seta diferenca"
    Set m = d.SelectSingleNode("//u:menu[@id='mnuArrows']/u:button[@id='btnDeltaFmt']")
    AssertTrue Not m Is Nothing
    AssertEqual "Delta: valor / % / ambos", AttrText(m, "label"): AssertEqual "delta", AttrText(m, "tag"): AssertEqual "Rx_OnDeltaFormat", AttrText(m, "onAction")
    AssertEqual "Rx_GetEnabled", AttrText(m, "getEnabled")
    AssertEqual "btnDeltaFmt", AttrText(d.SelectSingleNode("//u:button[@id='btnDelta']/following-sibling::*[1]"), "id"), "o formato logo depois do Delta"
    AssertTrue InStr(AttrText(n, "supertip"), "Esc") > 0, "supertip ensina a cancelar"
    AssertTrue InStr(AttrText(m, "supertip"), "selecionada") > 0, "supertip do formato diz o que ele altera"
    AssertEqual "tqq_arrows16", AttrText(n, "image"): AssertEqual "tqq_arrows16", AttrText(m, "image")
    AssertTrue InStr(AttrText(d.SelectSingleNode("//u:menu[@id='mnuArrows']"), "supertip"), "delta entre duas barras escolhidas com o mouse") > 0, "o supertip do menu cita o delta"
    AssertEqual 3, d.SelectNodes("//u:button[@onAction='Rx_OnAnnotation']").Length, "as 3 setas antigas continuam"
End Sub

' Os dois botoes do delta chegam ao core pelo caminho que avisa 'atualize' quando o core e anterior (RunCoreFeature) e nunca propagam erro.
Public Sub Test_Ribbon_Delta_Buttons_ReachCore()
    Dim code As String, en As Long, ed As String
    On Error GoTo Cleanup
    code = RibbonModuleCode()
    AssertTrue InStr(code, "RunCoreFeature ""Core_DeltaPick""") > 0, "Rx_OnDelta chama Core_DeltaPick"
    AssertTrue InStr(code, "RunCoreFeature ""Core_DeltaCycleFormat""") > 0, "Rx_OnDeltaFormat chama Core_DeltaCycleFormat"
    If Loader_CoreWorkbook() Is Nothing Then Loader_Log "TestRibbon: core nao carregado (parte do aviso pulada)": GoTo Cleanup
    Rx_TestSetFeatureProcSuffix "_inexistente"
    Loader_LastMsg = ""
    Rx_OnDelta Nothing
    AssertEqual RX_MSG_UPDATE, Loader_LastMsg, "Delta clicando nas barras com core anterior"
    Loader_LastMsg = ""
    Rx_OnDeltaFormat Nothing
    AssertEqual RX_MSG_UPDATE, Loader_LastMsg, "Delta: valor / % / ambos com core anterior"
    Rx_TestSetFeatureProcSuffix ""
    Err.Clear
Cleanup:
    en = Err.Number: ed = Err.Description
    Rx_TestSetFeatureProcSuffix ""
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

' Rotulos, tags e textos de ajuda exatos dos controles da Fase 3 (o tutorial e o Test-Docs dependem dos rotulos).
Public Sub Test_Ribbon_Phase3_LabelsAndTips()
    Dim d As Object, spec As Variant, i As Long, p() As String, n As Object
    Set d = LoadRibbon()
    spec = Array("btnButterfly|Borboleta|butterfly", "btnScatter|Dispersão|scatter", "btnBubble|Bolhas|bubble", _
                 "tglRoundSum|Rótulos com soma consistente|roundsum", "tglTrend|Linha de tendência|trend", "tglQuad|Quadrantes (médias)|quad", _
                 "btnBreak|Corte de eixo...|break", "btnBreakOff|Remover corte de eixo|breakoff", "btnRoundRange|Arredondar intervalo...|", _
                 "mnuPpt|PowerPoint|", "btnPptImage|Copiar como imagem|", "btnPptChart|Copiar como gráfico vinculado|")
    For i = LBound(spec) To UBound(spec)
        p = Split(spec(i), "|")
        Set n = d.SelectSingleNode("//*[@id='" & p(0) & "']")
        AssertTrue Not n Is Nothing, "controle ausente: " & p(0)
        AssertEqual p(1), AttrText(n, "label"), "rotulo de " & p(0)
        AssertEqual p(2), AttrText(n, "tag"), "tag de " & p(0)
        If p(0) <> "btnBreakOff" And p(0) <> "mnuPpt" Then AssertTrue Len(AttrText(n, "supertip")) > 20, "sem supertip: " & p(0)
    Next
    AssertEqual "PowerPoint", AttrText(d.SelectSingleNode("//*[@id='mnuPpt']"), "screentip")
    AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='btnSameScale']"), "supertip"), "corte de eixo") > 0, "Mesma escala: o supertip diz que grafico com corte de eixo nao entra"
    AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='tglRoundSum']"), "supertip"), "maior resto") > 0
End Sub

Public Sub Test_Ribbon_Toggles_And_Dropdowns_Phase2()
    Dim d As Object, n As Object, t As String
    Set d = LoadRibbon()
    AssertEqual 12, d.SelectNodes("//u:toggleButton").Length
    For Each n In d.SelectNodes("//u:toggleButton")
        t = n.Attributes.getNamedItem("tag").Text
        AssertTrue InStr(",totals,labels,connectors,ind100,vline,secondary,pieout,roundsum,trend,quad,pairdiff,pairabs,", "," & t & ",") > 0, "toggle invalido: " & t
        AssertEqual "Rx_OnToggle", AttrText(n, "onAction"), "toggle sem Rx_OnToggle: " & t
        AssertEqual "Rx_GetEnabled", n.Attributes.getNamedItem("getEnabled").Text, "toggle sem getEnabled: " & t
        AssertEqual "Rx_GetPressed", n.Attributes.getNamedItem("getPressed").Text
    Next
    AssertEqual 4, d.SelectNodes("//u:dropDown").Length
    AssertEqual "right,none,lineend,inbar", ItemIds(d.SelectSingleNode("//u:dropDown[@tag='legend']"))
    ' T11: so cagr, diff e bracket chegam ao core (Core_AddAnnotation ignora em silencio qualquer outra tag)
    AssertEqual 3, d.SelectNodes("//u:button[@onAction='Rx_OnAnnotation']").Length
    For Each n In d.SelectNodes("//u:button[@onAction='Rx_OnAnnotation']")
        t = n.Attributes.getNamedItem("tag").Text
        AssertTrue InStr(",cagr,diff,bracket,", "," & t & ",") > 0, "annotation kind invalido: " & t
    Next
End Sub

Public Sub Test_Ribbon_ImageMso_AreValid()
    Dim d As Object, n As Object, nm As String, pic As Object, checked As Long
    Set d = LoadRibbon()
    For Each n In d.SelectNodes("//*[@imageMso]")
        nm = n.Attributes.getNamedItem("imageMso").Text
        Set pic = Nothing
        On Error Resume Next
        Set pic = Application.CommandBars.GetImageMso(nm, 16, 16)
        On Error GoTo 0
        AssertTrue Not pic Is Nothing, "imageMso invalido: " & nm
        checked = checked + 1
    Next
    AssertTrue checked >= 15
End Sub

Public Sub Test_Ribbon_CoreState_WithoutChart()
    AssertEqual False, CBool(Loader_RunCore("Core_HasChart"))
End Sub

Public Sub Test_Ribbon_TutorialPath_FallsBackToUrl()
    Dim p As String
    Loader_SetRootOverride TempDir()
    p = Loader_TutorialPath()
    Loader_SetRootOverride ""
    AssertTrue Transport_IsHttp(p), "sem arquivo local deve usar a URL"
End Sub

' Todo Public Sub Rx_* tem tratamento de erro antes da primeira linha executavel (nunca propaga para a faixa).
Public Sub Test_Ribbon_AllCallbacks_HaveErrorContainment()
    Dim lines() As String, i As Long, j As Long, ln As String, nm As String, n As Long, t As String
    lines = Split(Replace(RibbonModuleCode(), vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If Left$(ln, 15) = "Public Sub Rx_O" Or Left$(ln, 15) = "Public Sub Rx_G" Then
            nm = Mid$(ln, 12, InStr(ln, "(") - 12)
            For j = i + 1 To UBound(lines)
                t = Trim$(lines(j))
                If Len(t) > 0 And Left$(t, 1) <> "'" And Left$(t, 4) <> "Dim " Then
                    AssertTrue Left$(t, 8) = "On Error", "sem On Error antes da 1a linha executavel em " & nm
                    Exit For
                End If
            Next
            n = n + 1
        End If
    Next
    AssertTrue n >= 14, "poucos callbacks Rx_ encontrados: " & n
End Sub

' Atualizar (btnRefresh) e "Mais opcoes" (mnuMore: contem Mesma escala, Escala automatica e Arredondar intervalo, sempre ativos) ficam
' sempre habilitados de proposito; os demais controles do grupo "Grafico selecionado" dependem do grafico selecionado. Dentro de mnuMore,
' os 7 toggles e os 2 botoes de corte de eixo tem getEnabled; os 2 botoes de escala e o Arredondar intervalo (age sobre celulas) nao.
' O menu PowerPoint e "Atualizar graficos no PowerPoint" ficam sempre ativos; so os 2 itens de copiar exigem um grafico.
Public Sub Test_Ribbon_SelectedChartGroup_EnabledExceptRefresh()
    Dim d As Object, n As Object, c As Long, id As String
    Set d = LoadRibbon()
    For Each n In d.SelectNodes("//u:group[@id='grpChart']/*[not(self::u:separator)]")
        id = n.Attributes.getNamedItem("id").Text
        If id = "btnRefresh" Or id = "mnuMore" Or id = "mnuPpt" Or id = "btnStyles" Then
            AssertTrue n.Attributes.getNamedItem("getEnabled") Is Nothing, id & " deve ficar sempre habilitado"
        Else
            AssertTrue Not n.Attributes.getNamedItem("getEnabled") Is Nothing, "sem getEnabled: " & id
        End If
        c = c + 1
    Next
    AssertTrue c >= 10
    For Each n In d.SelectNodes("//u:menu[@id='mnuMore']/u:toggleButton")
        AssertTrue Not n.Attributes.getNamedItem("getEnabled") Is Nothing, "toggle de Mais opcoes sem getEnabled: " & NodeId(n)
    Next
    c = 0
    For Each n In d.SelectNodes("//u:menu[@id='mnuMore']/u:button")
        id = NodeId(n)
        Select Case id
            Case "btnBreak", "btnBreakOff"
                AssertEqual "Rx_GetEnabled", AttrText(n, "getEnabled"), "corte de eixo segue a matriz (Core_Supports): " & id
            Case "btnLinkTitle"
                AssertEqual "Rx_GetEnabled", AttrText(n, "getEnabled"), "titulo ligado exige um grafico ThinQQ"
                AssertEqual "", AttrText(n, "tag"), "titulo ligado vale para todo tipo (sem tag: so exige grafico)"
            Case "btnSameScale", "btnResetScale", "btnRoundRange"
                AssertTrue n.Attributes.getNamedItem("getEnabled") Is Nothing, "deve ficar sempre ativo: " & id
            Case Else
                AssertTrue False, "botao inesperado em Mais opcoes: " & id
        End Select
        c = c + 1
    Next
    AssertEqual 6, c, "botoes de Mais opcoes"
    ' Setas: cada seta consulta a matriz pela propria tag; "Remover setas" so exige um grafico (sem tag)
    For Each n In d.SelectNodes("//u:menu[@id='mnuArrows']/u:button")
        AssertTrue Not n.Attributes.getNamedItem("getEnabled") Is Nothing, "seta sem getEnabled: " & NodeId(n)
    Next
    ' PowerPoint: o menu e Atualizar graficos no PowerPoint ficam sempre ativos; copiar (imagem/grafico) so exige um grafico (sem tag)
    AssertTrue d.SelectSingleNode("//u:menu[@id='mnuPpt']/@getEnabled") Is Nothing, "o menu PowerPoint fica sempre ativo"
    For Each n In d.SelectNodes("//u:menu[@id='mnuPpt'] | //u:menu[@id='mnuPpt']/u:button")
        AssertEqual "", AttrText(n, "tag"), "PowerPoint nao consulta a matriz: " & NodeId(n)
        If NodeId(n) = "btnPptImage" Or NodeId(n) = "btnPptChart" Then AssertEqual "Rx_GetEnabled", AttrText(n, "getEnabled"), "copiar exige grafico: " & NodeId(n)
        If NodeId(n) = "btnPptSync" Then AssertTrue n.Attributes.getNamedItem("getEnabled") Is Nothing, "Atualizar graficos no PowerPoint: sempre ativo"
        If NodeId(n) = "btnPptLinkText" Then AssertTrue n.Attributes.getNamedItem("getEnabled") Is Nothing, "Vincular celula ao texto: age sobre a celula, sempre ativo"
        If NodeId(n) = "btnPptTable" Then AssertTrue n.Attributes.getNamedItem("getEnabled") Is Nothing, "Tabela vinculada: age sobre as celulas, sempre ativo"
    Next
End Sub

' Sem core com Core_InsertSamples (Tarefa 15) o botao nao pode bloquear nem propagar erro. Com o core, a pasta de exemplos aberta pelo teste
' e fechada no fim (diferenca entre as pastas antes e depois) e os observadores do core voltam ao estado de antes (Onda B: ligados ou nao, e
' o loader registrado): nada sobra nem falta no Excel dos testes seguintes.
Public Sub Test_Ribbon_OnSamples_WithoutCoreProc_DoesNotBlock()
    Dim n0 As Long, w As Workbook, found As Boolean, sh As Object, before As Object, en As Long, ed As String, wasActive As Boolean, ln As String
    On Error GoTo Cleanup
    If Not Loader_CoreWorkbook() Is Nothing Then
        wasActive = CBool(Loader_RunCore("Watchers_IsActive")): ln = CStr(Loader_RunCore("Watchers_LoaderName"))
    End If
    Set before = CreateObject("Scripting.Dictionary")
    For Each w In Application.Workbooks
        before(w.Name) = True
    Next
    n0 = Application.Workbooks.Count
    Loader_LastMsg = ""
    Rx_OnSamples Nothing
    AssertTrue InStr(Loader_LastMsg, "Não foi possível") = 0, "sem falha ao inserir os exemplos: " & Loader_LastMsg
    If Loader_CoreWorkbook() Is Nothing Then
        AssertTrue InStr(Loader_LastMsg, "não está carregado") > 0, "sem core: o aviso de core ausente (" & Loader_LastMsg & ")"
    Else
        AssertEqual n0 + 1, Application.Workbooks.Count, "com o core: a pasta de exemplos foi aberta"
        For Each w In Application.Workbooks
            For Each sh In w.Worksheets
                If sh.Name = "Exemplos" Then found = True
            Next
        Next
        AssertTrue found, "a pasta aberta tem a planilha Exemplos"
    End If
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Loader_RunCore "Watchers_Stop"
    CloseNewWorkbooks before
    If wasActive Then Loader_RunCore "Watchers_Start"            ' o estado de antes (o Stop acima desliga e esquece o loader)
    If Len(ln) > 0 Then Loader_RunCore "Watchers_SetLoader", ln
    On Error GoTo 0
    If en = 0 Then AssertEqual n0, Application.Workbooks.Count, "a pasta de exemplos foi fechada no fim do teste"
    If en = 0 And Not Loader_CoreWorkbook() Is Nothing Then
        AssertEqual wasActive, CBool(Loader_RunCore("Watchers_IsActive")), "observadores de volta ao estado de antes"
        AssertEqual ln, CStr(Loader_RunCore("Watchers_LoaderName")), "loader registrado de volta"
    End If
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

' Fecha sem salvar toda pasta aberta que nao estava em before (nomes); nunca propaga erro.
Private Sub CloseNewWorkbooks(before As Object)
    Dim i As Long, w As Workbook
    On Error Resume Next
    If before Is Nothing Then Exit Sub
    For i = Application.Workbooks.Count To 1 Step -1
        Set w = Application.Workbooks(i)
        If Not before.Exists(w.Name) Then
            w.Saved = True
            w.Close False
        End If
    Next
End Sub

Public Sub Test_Ribbon_TutorialUrl_IsGithubPages()
    AssertEqual "https://israeldias-hub.github.io/ThinQQ/tutorial/tutorial.html", TUTORIAL_URL
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Habilitacao por capacidade (usa o core de teste aberto junto com o loader)
' ---------------------------------------------------------------------------------------------------------------------

' Tres tabelas na mesma planilha, bem separadas: pizza (A1:B5), cascata de uma serie (A20:B23) e cascata de duas series (A40:C43).
Private Function CapabilityFixture(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:B1").Value = Array("Categoria", "Valor")
    ws.Range("A2:B2").Value = Array("A", 10): ws.Range("A3:B3").Value = Array("B", 20)
    ws.Range("A4:B4").Value = Array("C", 30): ws.Range("A5:B5").Value = Array("D", 40)
    ws.Range("A20:B20").Value = Array("Etapa", "Loja")
    ws.Range("A21:B21").Value = Array("Inicio", 100): ws.Range("A22:B22").Value = Array("Ajuste", -30): ws.Range("A23:B23").Value = Array("Fim", "e")
    ws.Range("A40:C40").Value = Array("Etapa", "Loja", "Online")
    ws.Range("A41:C41").Value = Array("Inicio", 100, 50): ws.Range("A42:C42").Value = Array("Ajuste", -30, 10): ws.Range("A43:C43").Value = Array("Fim", "e", "e")
    Set CapabilityFixture = ws
End Function

' Insere pelo core o grafico da tabela e o deixa selecionado; depois faz o que o evento de selecao do core faz
' (Watchers_NotifySelection -> Loader_InvalidateRibbon): o cache da faixa NAO e zerado diretamente pelo teste.
Private Sub InsertAndSelect(ws As Worksheet, addr As String, kind As String)
    ws.Range(addr).Select
    Loader_RunCore "Core_Insert", kind
    Loader_RunCore "Watchers_NotifySelection"
End Sub

Private Sub EnabledCleanup()
    On Error Resume Next
    Rx_TestSetSupportsProc ""
    Loader_RunCore "Watchers_Stop"
    Loader_InvalidateRibbon
End Sub

' A matriz de capacidades vale por grafico: pizza (Totais/Conectores/Espaco/Setas cinza; Rotulos fora ativo), cascata de uma serie (sem
' legenda) e de duas (com legenda). Cada troca de grafico passa pelo evento de selecao do core, que zera o cache do ciclo: sem isso o
' resultado do grafico anterior ficaria valendo. Sem tag basta haver grafico; celula selecionada: nada habilitado.
Public Sub Test_Ribbon_Enabled_FollowsCapabilityOfSelectedChart()
    Dim wb As Workbook, ws As Worksheet, en As Long, ed As String, t As Variant
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set ws = CapabilityFixture(wb)
    Loader_RunCore "Watchers_SetLoader", ThisWorkbook.Name
    ws.Range("A1").Select
    Loader_InvalidateRibbon
    For Each t In Array("totals", "labels", "pieout", "legend", "cagr", "bracket", "")
        AssertFalse Rx_IsEnabledFor(CStr(t)), "celula selecionada: '" & t & "' deveria estar desabilitado"
    Next
    InsertAndSelect ws, "A1:B5", "pie"
    For Each t In Array("totals", "connectors", "gap", "cagr", "diff", "bracket", "ind100", "vline", "secondary")
        AssertFalse Rx_IsEnabledFor(CStr(t)), "pizza: '" & t & "' deveria estar desabilitado"
    Next
    For Each t In Array("labels", "numfmt", "style", "legend", "pieout", "")
        AssertTrue Rx_IsEnabledFor(CStr(t)), "pizza: '" & t & "' deveria estar habilitado"
    Next
    For Each t In Array("roundsum", "trend", "quad", "break")
        AssertFalse Rx_IsEnabledFor(CStr(t)), "pizza: '" & t & "' deveria estar desabilitado"
    Next
    InsertAndSelect ws, "A20:B23", "waterfall"
    For Each t In Array("connectors", "gap", "cagr", "diff", "labels", "")
        AssertTrue Rx_IsEnabledFor(CStr(t)), "cascata: '" & t & "' deveria estar habilitado (cache da pizza nao pode valer)"
    Next
    For Each t In Array("totals", "pieout", "bracket", "ind100", "vline", "secondary", "legend")
        AssertFalse Rx_IsEnabledFor(CStr(t)), "cascata de uma serie: '" & t & "' deveria estar desabilitado"
    Next
    InsertAndSelect ws, "A40:C43", "waterfall"
    AssertTrue Rx_IsEnabledFor("legend"), "cascata com duas series tem legenda"
    AssertFalse Rx_IsEnabledFor("pieout")
    AssertFalse Rx_IsEnabledFor("pairdiff"), "cascata comum: sem setas de variacao"
    InsertAndSelect ws, "A40:C43", "waterfallpair"
    For Each t In Array("pairdiff", "pairabs", "inbar", "totals", "connectors", "labels", "gap", "legend", "numfmt", "")
        AssertTrue Rx_IsEnabledFor(CStr(t)), "cascata comparativa: '" & t & "' deveria estar habilitado"
    Next
    For Each t In Array("cagr", "diff", "bracket", "vline", "roundsum", "break", "pieout")
        AssertFalse Rx_IsEnabledFor(CStr(t)), "cascata comparativa: '" & t & "' deveria estar desabilitado"
    Next
    InsertAndSelect ws, "A1:B5", "clustered"
    AssertTrue Rx_IsEnabledFor("break"), "colunas agrupadas aceitam corte de eixo"
    AssertFalse Rx_IsEnabledFor("roundsum"), "colunas agrupadas nao tem rotulos de segmento com soma"
    AssertFalse Rx_IsEnabledFor("trend")
    InsertAndSelect ws, "A1:B5", "stacked"
    AssertTrue Rx_IsEnabledFor("roundsum"), "colunas empilhadas: soma consistente"
    AssertTrue Rx_IsEnabledFor("break"), "colunas empilhadas: corte de eixo"
    AssertFalse Rx_IsEnabledFor("breakoff"), "sem corte: Remover corte de eixo cinza"
    AssertFalse Rx_IsEnabledFor("quad")
    ws.Range("A1").Select
    Loader_RunCore "Watchers_NotifySelection"
    AssertFalse Rx_IsEnabledFor("legend"), "de volta a uma celula"
    AssertFalse Rx_IsEnabledFor("")
Cleanup:
    en = Err.Number: ed = Err.Description
    EnabledCleanup
    On Error Resume Next
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

' Core sem Core_Supports (versao 1.0.0): Loader_RunCore devolve Empty e a faixa volta ao comportamento da Fase 1 (habilitado quando ha grafico).
Public Sub Test_Ribbon_Enabled_FallsBackToHasChart_WhenCoreLacksSupports()
    Dim wb As Workbook, ws As Worksheet, en As Long, ed As String
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set ws = CapabilityFixture(wb)
    InsertAndSelect ws, "A1:B5", "pie"
    Loader_InvalidateRibbon                ' o cache do ciclo pode vir de outro teste (o evento de selecao so zera se o core conhece o loader)
    AssertFalse Rx_IsEnabledFor("totals"), "pre-condicao: o core novo recusa Totais na pizza"
    Rx_TestSetSupportsProc "Core_SupportsNaoExiste"
    Loader_InvalidateRibbon
    AssertTrue Rx_IsEnabledFor("totals"), "core antigo: com grafico, habilitado como na Fase 1"
    AssertTrue Rx_IsEnabledFor("pieout")
    AssertTrue Rx_IsEnabledFor("")
    ws.Range("A1").Select
    Loader_InvalidateRibbon
    AssertFalse Rx_IsEnabledFor("totals"), "core antigo, sem grafico: desabilitado"
    AssertFalse Rx_IsEnabledFor("")
Cleanup:
    en = Err.Number: ed = Err.Description
    EnabledCleanup
    On Error Resume Next
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

' Mapeamento reverso (opcao gravada no grafico -> item do dropdown) com um grafico REAL, nao so os padroes sem grafico: colunas agrupadas
' nascem com Espaco 80 e legenda a direita, e o item de Espaco e "gap80"; escolher outro item (id -> valor do core, como Rx_OnDropdown faz)
' volta como o mesmo id.
Public Sub Test_Ribbon_SelectedId_ReverseMapping_WithRealChart()
    Dim wb As Workbook, ws As Worksheet, en As Long, ed As String
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set ws = CapabilityFixture(wb)
    InsertAndSelect ws, "A1:B5", "clustered"
    AssertEqual "gap80", Rx_SelectedIdFor("gap"), "colunas agrupadas nascem com Espaco 80: item gap80"
    AssertEqual "right", Rx_SelectedIdFor("legend"), "legenda a direita"
    AssertEqual "int", Rx_SelectedIdFor("numfmt")
    Loader_RunCore "Core_SetOption", "gap", Rx_OptionValueFor("gap", "gap120")
    AssertEqual "gap120", Rx_SelectedIdFor("gap"), "item escolhido volta como o mesmo id"
    Loader_RunCore "Core_SetOption", "legend", "none"
    AssertEqual "none", Rx_SelectedIdFor("legend")
    Loader_RunCore "Core_SetOption", "numfmt", "pct0"
    AssertEqual "pct0", Rx_SelectedIdFor("numfmt")
    Loader_RunCore "Core_SetOption", "palette", "cinza"
    AssertEqual "Tons de cinza", CStr(Loader_RunCore("Core_GetStyleName")), "a paleta antiga continua aceita pelo core e aparece como o estilo embutido"
    AssertEqual "st1", GetStyleSelForTest(), "item do dropdown Estilo = Tons de cinza"
Cleanup:
    en = Err.Number: ed = Err.Description
    EnabledCleanup
    On Error Resume Next
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

' Sem grafico (ou sem opcao gravada) cada lista mostra o seu padrao; a legenda comeca em "Direita".
Public Sub Test_Ribbon_SelectedId_DefaultsWithoutChart()
    Dim wb As Workbook
    Set wb = NewTempWorkbook()
    AssertEqual "right", Rx_SelectedIdFor("legend")
    AssertEqual "int", Rx_SelectedIdFor("numfmt")
    AssertEqual "gap50", Rx_SelectedIdFor("gap"), "o item de Espaco e 'gap' + valor"
    AssertEqual "120", Rx_OptionValueFor("gap", "gap120"), "o core recebe so o numero"
    AssertEqual "dec1", Rx_OptionValueFor("numfmt", "dec1"), "os demais dropdowns usam o proprio id"
    CloseTemp wb
End Sub

' Mesma escala e Escala automatica chegam ao core mesmo sem grafico (o core avisa o que fez; no build de teste o aviso so e registrado).
Public Sub Test_Ribbon_ScaleButtons_ReachCore_WithoutBlocking()
    Dim wb As Workbook
    Set wb = NewTempWorkbook()
    Loader_RunCore "Core_ClearLastNotice"
    Rx_OnSameScale Nothing
    AssertTrue InStr(CStr(Loader_RunCore("Core_LastNotice")), "Mesma escala") > 0, "Core_SameScale nao foi chamado: " & CStr(Loader_RunCore("Core_LastNotice"))
    Loader_RunCore "Core_ClearLastNotice"
    Rx_OnResetScale Nothing
    AssertTrue InStr(CStr(Loader_RunCore("Core_LastNotice")), "Escala automática") > 0, "Core_ResetScale nao foi chamado: " & CStr(Loader_RunCore("Core_LastNotice"))
    CloseTemp wb
End Sub

' Corte de eixo pela faixa: Rx_OnAxisBreak chega a Core_SetAxisBreak (a resposta do InputBox vem do gancho do core no build Test) e grava o
' corte; Rx_OnAxisBreakOff chega a Core_RemoveAxisBreak e o remove.
Public Sub Test_Ribbon_AxisBreakButtons_ReachCore()
    Dim wb As Workbook, ws As Worksheet, en As Long, ed As String
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set ws = CapabilityFixture(wb)
    InsertAndSelect ws, "A1:B5", "stacked"
    AssertEqual "", CoreOpt("vbreak"), "pre-condicao: sem corte"
    Loader_RunCore "Core_TestSetAxisBreakAnswer", 25
    Rx_OnAxisBreak Nothing
    AssertTrue Len(CoreOpt("vbreak")) > 0, "Corte de eixo... nao chegou ao core"
    AssertEqual 25, Val(CoreOpt("vbreak")), "limite do corte"
    Rx_OnAxisBreakOff Nothing
    AssertEqual "", CoreOpt("vbreak"), "Remover corte de eixo nao chegou ao core"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    Loader_RunCore "Core_TestSetAxisBreakAnswer", Empty
    EnabledCleanup
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

' Arredondar intervalo pela faixa: sempre ativo (age sobre celulas, sem grafico) e chega a Core_RoundRange (respostas pelo gancho do core).
Public Sub Test_Ribbon_RoundRangeButton_ReachesCore()
    Dim wb As Workbook, ws As Worksheet, en As Long, ed As String
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set ws = CapabilityFixture(wb)
    ws.Range("B2:B5").Select
    Loader_RunCore "Core_TestSetRoundRangeAnswers", 0, ws.Range("E2")
    Rx_OnRoundRange Nothing
    AssertTrue ws.Range("E2").HasFormula, "Arredondar intervalo... nao gravou a formula (nao chegou ao core)"
    AssertTrue ws.Range("E5").HasFormula
    AssertEqual 100, Application.WorksheetFunction.Sum(ws.Range("E2:E5")), "soma dos arredondados"
Cleanup:
    en = Err.Number: ed = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

' PowerPoint pela faixa: "Copiar como grafico vinculado" chega a Core_CopyToPowerPoint("chart"), que recusa a pasta nunca salva ANTES de
' abrir o PowerPoint (o teste nao abre o PowerPoint). "Copiar como imagem" abriria o PowerPoint: confere-se o modo no codigo do callback.
Public Sub Test_Ribbon_PptButtons_ReachCore()
    Dim wb As Workbook, ws As Worksheet, en As Long, ed As String, code As String
    On Error GoTo Cleanup
    Set wb = NewTempWorkbook(): Set ws = CapabilityFixture(wb)
    InsertAndSelect ws, "A1:B5", "clustered"
    Loader_InvalidateRibbon                ' o cache do ciclo pode vir de outro teste
    AssertTrue Rx_IsEnabledFor(""), "PowerPoint: ativo com grafico selecionado"
    Loader_RunCore "Core_ClearLastNotice"
    Rx_OnPptChart Nothing
    AssertTrue InStr(CStr(Loader_RunCore("Core_LastNotice")), "Salve a pasta de trabalho") > 0, "Copiar como grafico vinculado nao chegou ao core: " & CStr(Loader_RunCore("Core_LastNotice"))
    code = RibbonModuleCode()
    AssertTrue InStr(code, "RunCoreFeature ""Core_CopyToPowerPoint"", ""image""") > 0, "Rx_OnPptImage passa o modo image"
    AssertTrue InStr(code, "RunCoreFeature ""Core_CopyToPowerPoint"", ""chart""") > 0, "Rx_OnPptChart passa o modo chart"
    AssertTrue InStr(code, "RunCoreFeature ""Core_PptSyncOpen""") > 0, "Rx_OnPptSync chama Core_PptSyncOpen (sem exigir grafico)"
    ws.Range("A1").Select
    Loader_InvalidateRibbon
    AssertFalse Rx_IsEnabledFor(""), "PowerPoint: cinza sem grafico"
Cleanup:
    en = Err.Number: ed = Err.Description
    EnabledCleanup
    On Error Resume Next
    If Not wb Is Nothing Then CloseTemp wb
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

Private Function CoreOpt(key As String) As String
    Dim v As Variant
    v = Loader_RunCore("Core_GetOption", key)
    If Not IsEmpty(v) Then CoreOpt = CStr(v)
End Function

' Icones proprios (Fase 3): nomes image= no padrao tqq_<nome>[16] e nunca junto de imageMso no mesmo controle (Tarefa 8 exige >= 13).
Public Sub Test_Ribbon_Images_NamesAndNoDoubleIcon()
    Dim d As Object, n As Object, re As Object, nm As String, c As Long
    Set d = LoadRibbon()
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^tqq_[a-z0-9_]+$"
    For Each n In d.SelectNodes("//*[@image]")
        nm = n.Attributes.getNamedItem("image").Text
        AssertTrue re.Test(nm), "nome de imagem invalido: " & nm
        ' a mensagem so e montada na falha e tolera controle sem id (NodeId nunca da erro 91)
        If Not n.Attributes.getNamedItem("imageMso") Is Nothing Then AssertTrue False, "image e imageMso no mesmo controle: " & NodeId(n)
        c = c + 1
    Next
    AssertTrue c >= 13, "icones proprios: " & c
    AssertEqual 0, d.SelectNodes("//*[@image and @imageMso]").Length, "nenhum controle com image e imageMso"
End Sub

' As reservas imageMso da tabela de icones (build\lib\Ribbon-Icons.ps1, usadas no modo -RibbonIcons Mso) existem neste Office. O teste
' Test_Ribbon_ImageMso_AreValid le o XML fonte e nunca ve as reservas.
Public Sub Test_Ribbon_IconFallbackMso_AreValid()
    Dim txt As String, re As Object, m As Object, pic As Object, checked As Long
    txt = Transport_ReadTextFile(RepoPath() & "\build\lib\Ribbon-Icons.ps1")
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "Mso = '([^']+)'"
    For Each m In re.Execute(txt)
        Set pic = Nothing
        On Error Resume Next
        Set pic = Application.CommandBars.GetImageMso(CStr(m.SubMatches(0)), 16, 16)
        On Error GoTo 0
        AssertTrue Not pic Is Nothing, "imageMso de reserva invalido: " & m.SubMatches(0)
        checked = checked + 1
    Next
    AssertEqual 16, checked, "uma reserva por icone da tabela (1.11.0: + tqq_flag e tqq_links)"
End Sub

' Onda B (B3.10): loader novo com core anterior (sem os procedimentos novos): Loader_RunCore marca o procedimento inexistente e os botoes
' sempre ativos (Arredondar intervalo, PowerPoint) avisam que e preciso atualizar em vez de nao fazer nada. Com o core atual: sem o aviso.
Public Sub Test_Ribbon_FeatureWithOlderCore_AsksToUpdate()
    Dim en As Long, ed As String
    On Error GoTo Cleanup
    If Loader_CoreWorkbook() Is Nothing Then Loader_Log "TestRibbon: core nao carregado (pulado)": Exit Sub
    Loader_RunCore "Core_NaoExisteOndaB"
    AssertTrue Loader_RunCoreMissed(), "procedimento inexistente no core"
    Loader_RunCore "Core_Version"
    AssertFalse Loader_RunCoreMissed(), "a chamada seguinte zera a marca"
    Rx_TestSetFeatureProcSuffix "_inexistente"
    Loader_LastMsg = ""
    Rx_OnRoundRange Nothing
    AssertEqual RX_MSG_UPDATE, Loader_LastMsg, "Arredondar intervalo"
    Loader_LastMsg = ""
    Rx_OnPptImage Nothing
    AssertEqual RX_MSG_UPDATE, Loader_LastMsg, "PowerPoint (imagem)"
    Loader_LastMsg = ""
    Rx_OnPptChart Nothing
    AssertEqual RX_MSG_UPDATE, Loader_LastMsg, "PowerPoint (grafico vinculado)"
    Loader_LastMsg = ""
    Rx_OnPptSync Nothing
    AssertEqual RX_MSG_UPDATE, Loader_LastMsg, "PowerPoint (atualizar graficos)"
    Rx_TestSetFeatureProcSuffix ""
    Loader_LastMsg = ""
    Rx_OnPptImage Nothing                                       ' core atual, sem grafico ativo: nada (o core decide), sem o aviso
    AssertTrue Loader_LastMsg <> RX_MSG_UPDATE, "core atual: sem o aviso de atualizar"
    Err.Clear                                                   ' os botoes usam On Error Resume Next: um erro ja tratado no core nao e falha aqui
Cleanup:
    en = Err.Number: ed = Err.Description
    Rx_TestSetFeatureProcSuffix ""
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

Public Sub Test_Ribbon_PptSync_LabelAndTip()
    Dim d As Object, n As Object
    Set d = LoadRibbon()
    Set n = d.SelectSingleNode("//*[@id='btnPptSync']")
    AssertTrue Not n Is Nothing
    AssertEqual "Atualizar gráficos no PowerPoint", AttrText(n, "label")
    AssertEqual "Refresh", AttrText(n, "imageMso")
    AssertTrue InStr(AttrText(n, "supertip"), "Vínculos") > 0, "o supertip aponta para o dialogo Vinculos"
End Sub

' Atualizacao automatica (1.6.0): duas caixas de selecao no menu PowerPoint, depois do Atualizar graficos e de um separador.
Public Sub Test_Ribbon_PptAutoCheckBoxes()
    Dim d As Object
    Set d = LoadRibbon()
    AssertTrue d.SelectSingleNode("//*[@id='chkPptAutoOn']") Is Nothing, "1.11.0: opcoes do automatico no dialogo Vinculos do PowerPoint"
    AssertTrue d.SelectSingleNode("//*[@id='chkPptAutoNew']") Is Nothing
    AssertTrue d.SelectSingleNode("//*[@id='msepPptAuto']") Is Nothing
    AssertEqual "btnPptFrames", AttrText(d.SelectSingleNode("//u:button[@id='btnPptSync']/following-sibling::*[1]"), "id")
    AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='btnPptSync']"), "supertip"), "Vínculos") > 0, "aponta para o dialogo do PowerPoint"
    AssertEqual 0, InStr(AttrText(d.SelectSingleNode("//*[@id='btnPptSync']"), "supertip"), "nada muda sozinho"), "sem prometer que nada muda sozinho (revisao final I3)"
End Sub

' Rotulos exatos (o tutorial os cita) e supertips que dizem o que a opcao faz de verdade.
Public Sub Test_Ribbon_PptAutoCheckBoxes_LabelsAndTips()
    Dim d As Object
    Set d = LoadRibbon()
    AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='mnuPpt']"), "supertip"), "controla a atualização automática") = 0, "o menu nao controla mais o automatico"
    AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='mnuPpt']"), "supertip"), "Vínculos") > 0
End Sub

' Core sem os procedimentos (sufixo de teste): a caixa fica desmarcada e o clique avisa "Atualize".
Public Sub Test_Ribbon_PptFlag_OldCore()
    Dim en As Long, ed As String
    On Error GoTo Cleanup
    Rx_TestSetFeatureProcSuffix "_naoexiste"
    AssertFalse Rx_PptFlagValue("autoon"), "core sem Core_PptGetFlag: desmarcada, sem erro"
    AssertFalse Rx_PptFlagValue("autonew"), "core sem Core_PptGetFlag: desmarcada, sem erro"
    Loader_LastMsg = ""
    Rx_OnPptFlagCore "autoon", True
    AssertEqual RX_MSG_UPDATE, Loader_LastMsg
Cleanup:
    en = Err.Number: ed = Err.Description
    Rx_TestSetFeatureProcSuffix ""
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

' Estado real: a caixa le o que o core guarda (padrao: automatico ligado, novos vinculos manuais) e o clique grava os dois valores
' (inclusive "True"/"False" como 2o argumento: RunCoreFeature passa dois argumentos). Restaura os padroes no fim.
Public Sub Test_Ribbon_PptFlag_ReadsAndWritesCoreState()
    Dim en As Long, ed As String
    On Error GoTo Cleanup
    If Loader_CoreWorkbook() Is Nothing Then Loader_Log "TestRibbon: core nao carregado (pulado)": Exit Sub
    Loader_RunCore "Core_PptSetFlag", "autoon", True
    Loader_RunCore "Core_PptSetFlag", "autonew", False
    AssertTrue Rx_PptFlagValue("autoon"), "padrao: automatico ligado"
    AssertFalse Rx_PptFlagValue("autonew"), "padrao: novos vinculos manuais"
    Loader_LastMsg = ""
    Rx_OnPptFlagCore "autoon", False
    AssertFalse Rx_PptFlagValue("autoon"), "desligar grava"
    AssertTrue Rx_PptFlagValue("autonew") = False, "a outra opcao nao muda"
    Rx_OnPptFlagCore "autonew", True
    AssertTrue Rx_PptFlagValue("autonew"), "ligar grava"
    AssertFalse Rx_PptFlagValue("autoon"), "a outra opcao nao muda"
    Rx_OnPptFlagCore "autoon", True
    AssertTrue Rx_PptFlagValue("autoon")
    AssertTrue Loader_LastMsg <> RX_MSG_UPDATE, "core atual: sem o aviso de atualizar"
    AssertFalse Rx_PptFlagValue("tagqualquer"), "tag desconhecida: desmarcada"
Cleanup:
    en = Err.Number: ed = Err.Description
    Loader_RunCore "Core_PptSetFlag", "autoon", True
    Loader_RunCore "Core_PptSetFlag", "autonew", False
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

' ---- Parte D2 (1.7.0): Inserir tambem no slide e Mostrar/ocultar molduras ----

Public Sub Test_Ribbon_PptInsertAndFrames()
    Dim d As Object
    Set d = LoadRibbon()
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:checkBox[@id='chkPptInsert'][@getPressed='Rx_GetPptInsert'][@onAction='Rx_OnPptInsert']") Is Nothing
    AssertEqual "chkPptInsert", AttrText(d.SelectSingleNode("//u:menu[@id='mnuPpt']/*[1]"), "id"), "primeiro item do menu"
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:button[@id='btnPptFrames'][@onAction='Rx_OnPptFrames']") Is Nothing
    AssertTrue d.SelectSingleNode("//u:button[@id='btnPptFrames']/@getEnabled") Is Nothing, "sempre ativo"
    AssertEqual "Inserir também no slide", AttrText(d.SelectSingleNode("//*[@id='chkPptInsert']"), "label")
    AssertEqual "Mostrar/ocultar molduras", AttrText(d.SelectSingleNode("//*[@id='btnPptFrames']"), "label")
End Sub

Public Sub Test_Ribbon_OnInsert_RoutesToPpt()
    Dim en As Long, ed As String
    On Error GoTo Cleanup
    Rx_SetPptInsertForTest True
    AssertEqual "Core_InsertToPowerPoint", Rx_InsertProcFor()
    Rx_SetPptInsertForTest False
    AssertEqual "Core_Insert", Rx_InsertProcFor()
Cleanup:
    en = Err.Number: ed = Err.Description
    Rx_SetPptInsertForTest False
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

' ---- Parte E (1.9.0): Vincular celula ao texto selecionado (menu PowerPoint) e Titulo ligado a uma celula (Mais opcoes) ----

Public Sub Test_Ribbon_LinkTextAndTitle()
    Dim d As Object, code As String, en As Long, ed As String
    On Error GoTo Cleanup
    Set d = LoadRibbon()
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:button[@id='btnPptLinkText'][@onAction='Rx_OnPptLinkText'][@image='tqq_ppt16']") Is Nothing
    AssertEqual "Vincular célula ao texto selecionado", AttrText(d.SelectSingleNode("//*[@id='btnPptLinkText']"), "label")
    AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='btnPptLinkText']"), "supertip"), "Texto vinculado ao Excel") > 0, "cita o botao do PowerPoint"
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnLinkTitle'][@onAction='Rx_OnLinkTitle'][@getEnabled='Rx_GetEnabled'][@image='tqq_options16']") Is Nothing
    AssertEqual "Título ligado a uma célula...", AttrText(d.SelectSingleNode("//*[@id='btnLinkTitle']"), "label")
    AssertEqual "btnLinkTitle", AttrText(d.SelectSingleNode("//u:button[@id='btnResetScale']/following-sibling::*[1]"), "id"), "depois de Escala automatica"
    code = RibbonModuleCode()
    AssertTrue InStr(code, "RunCoreFeature ""Core_PptLinkTextFromPpt""") > 0, "Rx_OnPptLinkText chama Core_PptLinkTextFromPpt"
    AssertTrue InStr(code, "RunCoreFeature ""Core_LinkTitle""") > 0, "Rx_OnLinkTitle chama Core_LinkTitle"
    If Loader_CoreWorkbook() Is Nothing Then Loader_Log "TestRibbon: core nao carregado (parte do aviso pulada)": GoTo Cleanup
    Rx_TestSetFeatureProcSuffix "_inexistente"
    Loader_LastMsg = ""
    Rx_OnPptLinkText Nothing
    AssertEqual RX_MSG_UPDATE, Loader_LastMsg, "Vincular celula ao texto com core anterior"
    Loader_LastMsg = ""
    Rx_OnLinkTitle Nothing
    AssertEqual RX_MSG_UPDATE, Loader_LastMsg, "Titulo ligado com core anterior"
    Rx_TestSetFeatureProcSuffix ""
    Err.Clear
Cleanup:
    en = Err.Number: ed = Err.Description
    Rx_TestSetFeatureProcSuffix ""
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

' ---- Parte F (1.10.0): Inserir tabela vinculada no slide (menu PowerPoint) ----

Public Sub Test_Ribbon_PptTable()
    Dim d As Object, code As String, en As Long, ed As String
    On Error GoTo Cleanup
    Set d = LoadRibbon()
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuPpt']/u:button[@id='btnPptTable'][@onAction='Rx_OnPptTable'][@image='tqq_ppt16']") Is Nothing
    AssertTrue d.SelectSingleNode("//u:button[@id='btnPptTable']/@getEnabled") Is Nothing, "sempre ativo (age sobre as celulas)"
    AssertEqual "Inserir tabela vinculada no slide", AttrText(d.SelectSingleNode("//*[@id='btnPptTable']"), "label")
    AssertEqual "Inserir tabela vinculada no slide", AttrText(d.SelectSingleNode("//*[@id='btnPptTable']"), "screentip")
    AssertTrue InStr(AttrText(d.SelectSingleNode("//*[@id='btnPptTable']"), "supertip"), "75 linhas e 25 colunas") > 0, "cita os limites"
    code = RibbonModuleCode()
    AssertTrue InStr(code, "RunCoreFeature ""Core_PptInsertTable""") > 0, "Rx_OnPptTable chama Core_PptInsertTable"
    If Loader_CoreWorkbook() Is Nothing Then Loader_Log "TestRibbon: core nao carregado (parte do aviso pulada)": GoTo Cleanup
    Rx_TestSetFeatureProcSuffix "_inexistente"
    Loader_LastMsg = ""
    Rx_OnPptTable Nothing
    AssertEqual RX_MSG_UPDATE, Loader_LastMsg, "Tabela vinculada com core anterior"
    Rx_TestSetFeatureProcSuffix ""
    Err.Clear
Cleanup:
    en = Err.Number: ed = Err.Description
    Rx_TestSetFeatureProcSuffix ""
    If en <> 0 Then Err.Raise en, "TestRibbon", ed
End Sub

' ---- Estilos (1.12.0): lista Estilo e botao Estilos... na faixa
Public Sub Test_Ribbon_Styles_XmlAndNames()
    Dim d As Object, v As Variant
    Set d = LoadRibbon()
    AssertTrue Not d.SelectSingleNode("//u:dropDown[@id='ddStyle'][@tag='style'][@getItemCount='Rx_StyleCount'][@getItemLabel='Rx_StyleLabel'][@getItemID='Rx_StyleId'][@getSelectedItemID='Rx_GetStyleSel'][@onAction='Rx_OnStyle'][@getEnabled='Rx_GetEnabled']") Is Nothing, "dropdown Estilo"
    AssertTrue d.SelectSingleNode("//u:dropDown[@id='ddPalette']") Is Nothing, "Paleta virou Estilo"
    AssertTrue Not d.SelectSingleNode("//u:dropDown[@id='ddLegend']/following-sibling::*[1][self::u:button][@id='btnStyles'][@onAction='Rx_OnStyles']") Is Nothing, "btnStyles logo depois de ddLegend"
    v = Rx_StyleNamesFor("A|B|C")
    AssertEqual 3, UBound(v) - LBound(v) + 1
    v = Rx_StyleNamesFor("")
    AssertEqual 1, UBound(v) - LBound(v) + 1
    AssertEqual "Padrão ThinQQ", CStr(v(0))
End Sub
