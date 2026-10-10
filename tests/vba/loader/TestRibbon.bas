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
                  "Rx_OnSamples", "Rx_OnTutorial", "Rx_OnCheckUpdates", "Rx_OnSettings", "Rx_OnAbout")
    For i = LBound(names) To UBound(names)
        AssertTrue InStr(xml, names(i)) > 0, "callback ausente no XML: " & names(i)
    Next
    AssertTrue InStr(xml, "tag=""waterfall""") > 0
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

Public Sub Test_Ribbon_ControlIds_Unique()
    Dim d As Object, n As Object, seen As Object, id As String
    Set d = LoadRibbon()
    Set seen = CreateObject("Scripting.Dictionary")
    For Each n In d.SelectNodes("//*[@id]")
        id = n.Attributes.getNamedItem("id").Text
        AssertFalse seen.Exists(id), "id duplicado: " & id
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
        AssertTrue IsNumeric(v), "item de gap nao numerico: " & v
        AssertTrue CLng(v) >= 0 And CLng(v) <= 500, "gap fora de 0..500: " & v
    Next
    AssertEqual 3, d.SelectNodes("//u:dropDown").Length
End Sub

Public Sub Test_Ribbon_Tags_MatchCoreKeys()
    Dim d As Object, n As Object, t As String
    Set d = LoadRibbon()
    AssertEqual 5, d.SelectNodes("//u:button[@onAction='Rx_OnInsert']").Length
    For Each n In d.SelectNodes("//u:button[@onAction='Rx_OnInsert']")
        t = n.Attributes.getNamedItem("tag").Text
        AssertTrue InStr(",stacked,clustered,stacked100,waterfall,line,", "," & t & ",") > 0, "kind invalido: " & t
    Next
    AssertEqual 3, d.SelectNodes("//u:toggleButton").Length
    For Each n In d.SelectNodes("//u:toggleButton")
        t = n.Attributes.getNamedItem("tag").Text
        AssertTrue InStr(",totals,labels,connectors,", "," & t & ",") > 0, "toggle key invalida: " & t
    Next
    For Each n In d.SelectNodes("//u:button[@onAction='Rx_OnAnnotation']")
        t = n.Attributes.getNamedItem("tag").Text
        AssertTrue InStr(",cagr,diff,", "," & t & ",") > 0, "annotation kind invalido: " & t
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
