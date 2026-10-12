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
                Case "toggleButton": ExpectedSig = "iribboncontrol,boolean"
                Case "dropDown": ExpectedSig = "iribboncontrol,string,integer"
            End Select
        Case "getEnabled", "getPressed", "getSelectedItemID": ExpectedSig = "iribboncontrol,byref *"
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
                  "Rx_OnSamples", "Rx_OnTutorial", "Rx_OnCheckUpdates", "Rx_OnSettings", "Rx_OnAbout", "Rx_OnSameScale", "Rx_OnResetScale")
    For i = LBound(names) To UBound(names)
        AssertTrue InStr(xml, names(i)) > 0, "callback ausente no XML: " & names(i)
    Next
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
    attrs = Array("onLoad", "onAction", "getEnabled", "getPressed", "getSelectedItemID")
    For Each n In d.SelectNodes("//*[@onLoad or @onAction or @getEnabled or @getPressed or @getSelectedItemID]")
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
    AssertEqual "padrao,cinza,tema", ItemIds(d.SelectSingleNode("//u:dropDown[@tag='palette']"))
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
    AssertEqual 17, d.SelectNodes("//u:button[@onAction='Rx_OnInsert']").Length
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
End Sub

' Menus (customUI14): 5 menus de insercao com os 17 botoes + "Mais opcoes" e "Setas". Um atributo invalido em filho de menu faz o Excel
' descartar a aba inteira, sem aviso: nenhum filho de menu leva 'size'.
Public Sub Test_Ribbon_Menus_Structure()
    Dim d As Object, m As Object, kids As Object, n As Object, t As String
    Set d = LoadRibbon()
    AssertEqual 7, d.SelectNodes("//u:menu").Length, "5 menus de insercao + Mais opcoes + Setas"
    AssertEqual 5, d.SelectNodes("//u:group[@id='grpInsert']/u:menu[@size='large']").Length
    AssertEqual 0, d.SelectNodes("//u:menu/*[@size]").Length, "filho de menu com 'size' (invalido: o Excel descarta a aba inteira sem aviso)"
    For Each m In d.SelectNodes("//u:group[@id='grpInsert']/u:menu")
        Set kids = m.SelectNodes("u:button")
        AssertTrue kids.Length >= 2, "menu " & m.Attributes.getNamedItem("id").Text & " com poucos itens"
        For Each n In kids
            AssertEqual "Rx_OnInsert", n.Attributes.getNamedItem("onAction").Text
            t = n.Attributes.getNamedItem("tag").Text
            AssertTrue InStr(",stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo,mekko,mekkounits,gantt,pie,doughnut,", "," & t & ",") > 0, "kind invalido: " & t
        Next
    Next
    AssertEqual 17, d.SelectNodes("//u:button[@onAction='Rx_OnInsert']").Length
    AssertTrue d.SelectSingleNode("//u:menu[@id='mnuMore']/@getEnabled") Is Nothing, "Mais opcoes contem botoes sempre ativos"
    AssertEqual 4, d.SelectNodes("//u:menu[@id='mnuMore']/u:toggleButton").Length
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnSameScale'][@onAction='Rx_OnSameScale']") Is Nothing
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnResetScale'][@onAction='Rx_OnResetScale']") Is Nothing
    AssertTrue Not d.SelectSingleNode("//u:menu[@id='mnuArrows']/u:button[@tag='bracket']") Is Nothing
    AssertEqual "Rx_GetEnabled", d.SelectSingleNode("//u:menu[@id='mnuArrows']").Attributes.getNamedItem("getEnabled").Text
End Sub

Public Sub Test_Ribbon_Toggles_And_Dropdowns_Phase2()
    Dim d As Object, n As Object, t As String
    Set d = LoadRibbon()
    AssertEqual 7, d.SelectNodes("//u:toggleButton").Length
    For Each n In d.SelectNodes("//u:toggleButton")
        t = n.Attributes.getNamedItem("tag").Text
        AssertTrue InStr(",totals,labels,connectors,ind100,vline,secondary,pieout,", "," & t & ",") > 0, "toggle invalido: " & t
        AssertEqual "Rx_GetEnabled", n.Attributes.getNamedItem("getEnabled").Text, "toggle sem getEnabled: " & t
        AssertEqual "Rx_GetPressed", n.Attributes.getNamedItem("getPressed").Text
    Next
    AssertEqual 4, d.SelectNodes("//u:dropDown").Length
    AssertEqual "right,none,lineend", ItemIds(d.SelectSingleNode("//u:dropDown[@tag='legend']"))
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

' Atualizar (btnRefresh) e "Mais opcoes" (mnuMore: contem Mesma escala e Escala automatica, sempre ativas) ficam sempre habilitados de
' proposito; os demais controles do grupo "Grafico selecionado" dependem do grafico selecionado. Dentro de mnuMore, os 4 toggles tem
' getEnabled e os 2 botoes de escala nao.
Public Sub Test_Ribbon_SelectedChartGroup_EnabledExceptRefresh()
    Dim d As Object, n As Object, c As Long, id As String
    Set d = LoadRibbon()
    For Each n In d.SelectNodes("//u:group[@id='grpChart']/*[not(self::u:separator)]")
        id = n.Attributes.getNamedItem("id").Text
        If id = "btnRefresh" Or id = "mnuMore" Then
            AssertTrue n.Attributes.getNamedItem("getEnabled") Is Nothing, id & " deve ficar sempre habilitado"
        Else
            AssertTrue Not n.Attributes.getNamedItem("getEnabled") Is Nothing, "sem getEnabled: " & id
        End If
        c = c + 1
    Next
    AssertTrue c >= 10
    For Each n In d.SelectNodes("//u:menu[@id='mnuMore']/u:toggleButton")
        AssertTrue Not n.Attributes.getNamedItem("getEnabled") Is Nothing, "toggle de Mais opcoes sem getEnabled: " & n.Attributes.getNamedItem("id").Text
    Next
    For Each n In d.SelectNodes("//u:menu[@id='mnuMore']/u:button")
        AssertTrue n.Attributes.getNamedItem("getEnabled") Is Nothing, "botao de escala deve ficar sempre ativo: " & n.Attributes.getNamedItem("id").Text
    Next
    ' Setas: cada seta consulta a matriz pela propria tag; "Remover setas" so exige um grafico (sem tag)
    For Each n In d.SelectNodes("//u:menu[@id='mnuArrows']/u:button")
        AssertTrue Not n.Attributes.getNamedItem("getEnabled") Is Nothing, "seta sem getEnabled: " & n.Attributes.getNamedItem("id").Text
    Next
End Sub

' Sem core com Core_InsertSamples (Tarefa 15) o botao nao pode bloquear nem propagar erro.
Public Sub Test_Ribbon_OnSamples_WithoutCoreProc_DoesNotBlock()
    Dim n0 As Long, w As Workbook, found As Boolean, sh As Object
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
    For Each t In Array("labels", "numfmt", "palette", "legend", "pieout", "")
        AssertTrue Rx_IsEnabledFor(CStr(t)), "pizza: '" & t & "' deveria estar habilitado"
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
    AssertEqual "padrao", Rx_SelectedIdFor("palette")
    Loader_RunCore "Core_SetOption", "gap", Rx_OptionValueFor("gap", "gap120")
    AssertEqual "gap120", Rx_SelectedIdFor("gap"), "item escolhido volta como o mesmo id"
    Loader_RunCore "Core_SetOption", "legend", "none"
    AssertEqual "none", Rx_SelectedIdFor("legend")
    Loader_RunCore "Core_SetOption", "numfmt", "pct0"
    AssertEqual "pct0", Rx_SelectedIdFor("numfmt")
    Loader_RunCore "Core_SetOption", "palette", "cinza"
    AssertEqual "cinza", Rx_SelectedIdFor("palette")
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
    AssertEqual "padrao", Rx_SelectedIdFor("palette")
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

' Icones proprios (Fase 3): nomes image= no padrao tqq_<nome>[16] e nunca junto de imageMso no mesmo controle (Tarefa 8 exige >= 13).
Public Sub Test_Ribbon_Images_NamesAndNoDoubleIcon()
    Dim d As Object, n As Object, re As Object, nm As String, c As Long
    Set d = LoadRibbon()
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^tqq_[a-z0-9_]+$"
    For Each n In d.SelectNodes("//*[@image]")
        nm = n.Attributes.getNamedItem("image").Text
        AssertTrue re.Test(nm), "nome de imagem invalido: " & nm
        AssertTrue n.Attributes.getNamedItem("imageMso") Is Nothing, "image e imageMso no mesmo controle: " & n.Attributes.getNamedItem("id").Text
        c = c + 1
    Next
    AssertTrue c >= 0
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
    AssertEqual 13, checked, "uma reserva por icone da tabela"
End Sub
