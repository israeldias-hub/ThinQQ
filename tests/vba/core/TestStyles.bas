Attribute VB_Name = "TestStyles"
Option Explicit
' Estilos de grafico (1.12.0; spec 2026-10-23 secao 8.5): modelo puro, depositos, aplicar por papel e sobrevivencia.

Public Sub Test_StyleDef_Validate_Rules()
    AssertEqual "#70AD47", StyleDef_Validate("aumento", " #70ad47 ")
    AssertEqual "", StyleDef_Validate("aumento", "#70AD4"), "hex curto"
    AssertEqual "", StyleDef_Validate("aumento", "70AD47"), "sem #"
    AssertEqual "", StyleDef_Validate("aumento", "#GGGGGG")
    AssertEqual "auto", StyleDef_Validate("rotulo.cor", "AUTO")
    AssertEqual "#FFFFFF", StyleDef_Validate("rotulo.cor", "#ffffff")
    AssertEqual "nenhuma", StyleDef_Validate("primeira.borda", "nenhuma")
    AssertEqual "diagonal_larga_para_cima", StyleDef_Validate("primeira.padrao", "Diagonal_Larga_Para_Cima")
    AssertEqual "", StyleDef_Validate("primeira.padrao", "nenhum,xadrez_pequeno"), "virgula nunca casa dois itens"
    AssertEqual "", StyleDef_Validate("primeira.padrao", "qualquer")
    AssertEqual "9", StyleDef_Validate("fonte.tamanho", "9")
    AssertEqual "", StyleDef_Validate("fonte.tamanho", "5")
    AssertEqual "", StyleDef_Validate("fonte.tamanho", "9.5")
    AssertEqual "", StyleDef_Validate("espaco", "-1")
    AssertEqual "500", StyleDef_Validate("espaco", "500")
    AssertEqual "0.75", StyleDef_Validate("conector.espessura", "0,75")
    AssertEqual "", StyleDef_Validate("conector.espessura", "3.5")
    AssertEqual "", StyleDef_Validate("conector.espessura", "1.2.3")
    AssertEqual "pontos", StyleDef_Validate("conector.traco", "Pontos")
    AssertEqual "", StyleDef_Validate("conector.traco", "zigue")
    AssertEqual "1", StyleDef_Validate("rotulo.negrito", "1")
    AssertEqual "", StyleDef_Validate("rotulo.negrito", "sim")
    AssertEqual "Calibri Light", StyleDef_Validate("fonte", "Calibri Light")
    AssertEqual "", StyleDef_Validate("fonte", "Calibri;=x")
    AssertEqual "", StyleDef_Validate("desconhecida", "#000000")
    AssertEqual "", StyleDef_Validate("serie9", "#000000"), "so serie1..serie8"
End Sub

Public Sub Test_StyleDef_SafeName()
    AssertEqual "Cascata MC", StyleDef_SafeName("  Cascata MC  ")
    AssertEqual "ab", StyleDef_SafeName("a;=|""\/:*?<>&b")
    AssertEqual "Padrão ThinQQ", StyleDef_SafeName("Padrão ThinQQ")
    AssertEqual 40, Len(StyleDef_SafeName(String$(60, "x")))
    AssertEqual "", StyleDef_SafeName(vbTab & vbCr)
End Sub

Public Sub Test_StyleDef_ParseSerialize_RoundTrip()
    Dim d As Object, ign As Long, t As String, d2 As Object
    t = ChrW(65279) & STYLEDEF_HEADER & vbLf & "# comentario" & vbCrLf & "nome=Meu" & vbLf & "AUMENTO=#70ad47" & vbLf & "reducao=vermelho" & vbLf & "lixo" & vbLf & "outra=1" & vbLf & vbLf & "espaco=40"
    Set d = StyleDef_Parse(t, ign)
    AssertTrue Not d Is Nothing
    AssertEqual 3, ign, "reducao invalida, linha sem '=', chave desconhecida"
    AssertEqual "#70AD47", StyleDef_Get(d, "aumento")
    AssertEqual "40", StyleDef_Get(d, "espaco")
    AssertEqual "", StyleDef_Get(d, "reducao")
    t = StyleDef_Serialize(d)
    AssertEqual STYLEDEF_HEADER & vbCrLf & "nome=Meu" & vbCrLf & "aumento=#70AD47" & vbCrLf & "espaco=40", t, "ordem canonica"
    Set d2 = StyleDef_Parse(t, ign)
    AssertEqual 0, ign
    AssertEqual t, StyleDef_Serialize(d2), "ida e volta"
    AssertEqual 8, Len(StyleDef_Version(t))
    AssertTrue StyleDef_Version(t) <> StyleDef_Version(t & vbCrLf & "total=#000000")
End Sub

Public Sub Test_StyleDef_Parse_RefusesHostile()
    Dim ign As Long, i As Long, s As String
    AssertTrue StyleDef_Parse("", ign) Is Nothing
    AssertTrue StyleDef_Parse("nome=x", ign) Is Nothing, "sem cabecalho"
    AssertTrue StyleDef_Parse(STYLEDEF_HEADER & vbLf & String$(STYLEDEF_MAX_BYTES, "a"), ign) Is Nothing, "grande demais"
    s = STYLEDEF_HEADER
    For i = 1 To STYLEDEF_MAX_LINES + 1
        s = s & vbLf & "x=1"
    Next
    AssertTrue StyleDef_Parse(s, ign) Is Nothing, "linhas demais"
    AssertTrue Not StyleDef_Parse(STYLEDEF_HEADER & vbLf & "nome==<script>" & vbLf & "fonte=a&b", ign) Is Nothing
    AssertEqual 2, ign, "nome e fonte invalidos ignorados, nada executado"
End Sub

Public Sub Test_StyleDef_ColorsPatternsAndBuiltins()
    Dim n As Variant, d As Object, ign As Long
    AssertEqual RGB(112, 173, 71), StyleDef_ColorOf("#70AD47")
    AssertEqual -1, StyleDef_ColorOf("x")
    AssertEqual "#70AD47", StyleDef_HexOf(RGB(112, 173, 71))
    AssertEqual msoPatternWideUpwardDiagonal, StyleDef_PatternOf("diagonal_larga_para_cima")
    AssertEqual 0, StyleDef_PatternOf("nenhum")
    AssertEqual "diagonal_escura_para_cima", StyleDef_PatternName(msoPatternDarkUpwardDiagonal)
    AssertEqual "", StyleDef_PatternName(999)
    AssertEqual 5, UBound(StyleDef_BuiltinNames()) + 1
    For Each n In StyleDef_BuiltinNames()
        Set d = StyleDef_Parse(StyleDef_Builtin(CStr(n)), ign)
        AssertTrue Not d Is Nothing, CStr(n)
        AssertEqual 0, ign, "embutido sem linha invalida: " & CStr(n)
        AssertEqual CStr(n), StyleDef_Get(d, "nome")
        AssertTrue StyleDef_IsBuiltin(CStr(n))
    Next
    AssertEqual "padrao", StyleDef_BuiltinPalette("Padrão ThinQQ")
    AssertEqual "", StyleDef_BuiltinPalette("Cascata MC")
    Set d = StyleDef_Parse(StyleDef_Builtin("Cascata MC"), ign)
    AssertEqual "#70AD47", StyleDef_Get(d, "aumento")
    AssertEqual "#E8173C", StyleDef_Get(d, "reducao")
    AssertEqual "#D9D9D9", StyleDef_Get(d, "total")
    AssertEqual "diagonal_larga_para_cima", StyleDef_Get(d, "primeira.padrao")
    AssertEqual 3, StyleDef_SeriesCount(d)
    AssertFalse StyleDef_IsBuiltin("Meu")
End Sub

Private Function StyleTestDir() As String
    Dim d As String
    d = TempDir() & "\styles"
    If Len(Dir(d, vbDirectory)) = 0 Then MkDir d
    StyleStore_SetStoreDirForTest d
    StyleTestDir = d
End Function

Public Sub Test_StyleStore_UserSaveGetDelete_TestDirOnly()
    Dim d As String, t As String
    StyleStore_ResetTestHooks
    AssertEqual "", StyleStore_UserDir(), "build Test sem gancho: nunca a pasta real do usuario"
    AssertTrue Left$(StyleStore_SaveUser("X", STYLEDEF_HEADER), 1) = "!", "sem deposito: recusa"
    d = StyleTestDir()
    t = STYLEDEF_HEADER & vbCrLf & "nome=Meu Estilo" & vbCrLf & "aumento=#00FF00"
    AssertEqual "", StyleStore_SaveUser("Meu Estilo", t)
    AssertTrue Len(Dir(d & "\Meu Estilo" & STYLEDEF_EXT)) > 0
    AssertEqual t, StyleStore_UserText("meu estilo"), "sem diferenciar maiusculas"
    AssertTrue InStr("|" & StyleStore_Names() & "|", "|Meu Estilo|") > 0
    AssertTrue InStr("|" & StyleStore_Names() & "|", "|Cascata MC|") > 0, "embutidos na lista"
    AssertTrue Left$(StyleStore_SaveUser("Cascata MC", t), 1) = "!", "nome de embutido recusado"
    AssertEqual "", StyleStore_DeleteUser("Meu Estilo")
    AssertEqual "", StyleStore_UserText("Meu Estilo")
    StyleStore_ResetTestHooks
End Sub

Public Sub Test_StyleStore_Book_CustomXml_SaveReopen()
    Dim wb As Workbook, p As String, t As String
    Set wb = NewTempWorkbook()
    t = STYLEDEF_HEADER & vbCrLf & "nome=Da Pasta" & vbCrLf & "total=#123456"
    StyleStore_SaveBook wb, "Da Pasta", t
    StyleStore_SaveBook wb, "Outro <&>", STYLEDEF_HEADER
    AssertEqual t, Replace(StyleStore_BookText(wb, "da pasta"), vbLf, vbCrLf), "texto de volta (o XML normaliza CRLF)"
    AssertEqual 1, wb.CustomXMLParts.SelectByNamespace(STYLESTORE_NS).Count, "uma parte so"
    p = TempDir() & "\style-book.xlsx"
    wb.SaveAs p, xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Set wb = Workbooks.Open(p)
    AssertTrue InStr(StyleStore_BookText(wb, "Da Pasta"), "total=#123456") > 0, "sobrevive a salvar e reabrir"
    AssertEqual 2, StyleStore_BookAll(wb).Count
    AssertTrue InStr(StyleStore_Get("Da Pasta", wb), "#123456") > 0, "a pasta acha o estilo que o usuario nao tem"
    CloseTemp wb
    Kill p
End Sub

Public Sub Test_StyleStore_ImportExport_Validates()
    Dim d As String, p As String, nm As String, ign As Long, r As String, f As Integer
    StyleStore_ResetTestHooks
    d = StyleTestDir()
    p = d & "\entrada.tqqestilo"
    f = FreeFile
    Open p For Output As #f
    Print #f, STYLEDEF_HEADER
    Print #f, "nome=Recebido;x"
    Print #f, "aumento=#ABCDEF"
    Print #f, "comando=executar o calc"
    Close #f
    r = StyleStore_Import(p, nm, ign)
    AssertTrue Left$(r, 1) <> "!", r
    AssertEqual "entrada", nm, "nome invalido no arquivo: vale o nome do arquivo"
    AssertEqual 2, ign
    AssertTrue InStr(StyleStore_UserText("entrada"), "aumento=#ABCDEF") > 0
    AssertEqual "", StyleStore_Export("entrada", d & "\saida.tqqestilo")
    AssertTrue InStr(StyleStore_ReadFile(d & "\saida.tqqestilo", r), "aumento=#ABCDEF") > 0
    f = FreeFile
    Open d & "\ruim.tqqestilo" For Output As #f
    Print #f, "<xml>nada</xml>"
    Close #f
    AssertTrue Left$(StyleStore_Import(d & "\ruim.tqqestilo", nm, ign), 1) = "!", "sem cabecalho"
    StyleStore_ResetTestHooks
End Sub

' Tabela de cascata (a mesma regra de TestWaterfall): a 1a barra (2026) e um aumento a partir de zero; Base e um total; Perda reduz.
Private Function WfTable(ws As Worksheet) As Range
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("2026", 20)
    ws.Range("A3:B3").Value = Array("Ganho A", 0.6)
    ws.Range("A4:B4").Value = Array("Ganho B", 0.3)
    ws.Range("A5:B5").Value = Array("Base", "e")
    ws.Range("A6:B6").Value = Array("Perda", -0.9)
    Set WfTable = ws.Range("A1:B6")
End Function

Private Function NewWfChart(wb As Workbook) As Chart
    Dim rng As Range
    Set rng = WfTable(wb.Worksheets(1))
    rng.Select
    Core_Insert "waterfall"
    Set NewWfChart = Core_ActiveChart()
End Function

Private Function SeriesFill(c As Chart, nm As String) As Long
    SeriesFill = c.SeriesCollection(nm).Format.Fill.ForeColor.RGB
End Function

' Papeis da cascata: aumento, reducao, total; a 1a barra (aumento a partir de zero) hachurada com borda; zero de outra serie no ponto 1 nao
' ganha borda (linha fantasma na base).
Public Sub Test_StyleApply_Waterfall_Roles_AndHatch()
    Dim wb As Workbook, c As Chart, p As Point
    Set wb = NewTempWorkbook()
    Set c = NewWfChart(wb)
    Style_ApplyNamed c, "Cascata MC"
    AssertEqual "Cascata MC", Registry_GetOption(c, "style", "")
    AssertEqual 8, Len(Registry_GetOption(c, "stylever", ""))
    AssertEqual RGB(112, 173, 71), SeriesFill(c, "tqq_uppos")
    AssertEqual RGB(232, 23, 60), SeriesFill(c, "tqq_downneg")
    AssertEqual RGB(217, 217, 217), SeriesFill(c, "tqq_total")
    Set p = c.SeriesCollection("tqq_uppos").Points(1)
    AssertEqual msoFillPatterned, p.Format.Fill.Type, "1a barra hachurada"
    AssertEqual msoPatternWideUpwardDiagonal, p.Format.Fill.Pattern
    AssertEqual RGB(0, 0, 0), p.Format.Fill.ForeColor.RGB
    AssertEqual RGB(255, 255, 255), p.Format.Fill.BackColor.RGB
    AssertEqual msoTrue, p.Format.Line.Visible, "borda da 1a barra"
    AssertEqual msoFalse, c.SeriesCollection("tqq_total").Points(1).Format.Line.Visible, "total zero no ponto 1: sem borda"
    AssertTrue InStr(StyleStore_BookText(wb, "Cascata MC"), "aumento=#70AD47") > 0, "a pasta passa a levar o estilo"
    CloseTemp wb
End Sub

' Sobrevivencia (Review Focus 6): ligar/desligar conectores, mudar formato, Alterar dados com uma linha a mais, refresh com categoria nova.
Public Sub Test_StyleApply_SurvivesTogglesRebuildAndRefresh()
    Dim wb As Workbook, c As Chart, ws As Worksheet, i As Long, conn As Shape
    Set wb = NewTempWorkbook()
    Set c = NewWfChart(wb)
    Set ws = wb.Worksheets(1)
    Style_ApplyNamed c, "Cascata MC"
    Core_SetToggle "connectors", False
    Core_SetToggle "connectors", True
    For i = 1 To c.Shapes.Count
        If Left$(c.Shapes(i).Name, Len(ANN_PREFIX & "conn_")) = ANN_PREFIX & "conn_" Then
            Set conn = c.Shapes(i)
            Exit For
        End If
    Next
    AssertTrue Not conn Is Nothing, "conectores recriados"
    AssertEqual RGB(127, 127, 127), conn.Line.ForeColor.RGB
    AssertEqual msoLineRoundDot, conn.Line.DashStyle, "conector pontilhado reaplicado"
    Core_SetOption "numfmt", "dec1"
    AssertEqual RGB(112, 173, 71), SeriesFill(c, "tqq_uppos"), "trocar o formato nao perde as cores"
    ws.Range("A7:B7").Value = Array("Final", "e")
    AssertTrue Core_RebuildFromRange(c, ws.Range("A1:B7"))
    Set c = Core_ActiveChart()
    AssertEqual "Cascata MC", Registry_GetOption(c, "style", "")
    AssertEqual RGB(112, 173, 71), SeriesFill(c, "tqq_uppos"), "Alterar dados reaplica"
    AssertEqual msoFillPatterned, c.SeriesCollection("tqq_uppos").Points(1).Format.Fill.Type
    AssertEqual Style_Stamp(c), Registry_GetOption(c, "stylestamp", "")
    Core_SetOption "palette", "cinza"
    AssertEqual "", Registry_GetOption(c, "style", ""), "escolher uma paleta embutida tira o estilo"
    AssertTrue c.SeriesCollection("tqq_uppos").Points(1).Format.Fill.Type <> msoFillPatterned, "sem estilo: a hachura sai"
    AssertEqual msoFalse, c.SeriesCollection("tqq_uppos").Points(1).Format.Line.Visible, "sem estilo: a borda da 1a barra sai"
    CloseTemp wb
End Sub

' Sondas PX1 como teste: mudar um valor mantem a hachura (o carimbo nao muda: nada reaplicado); a 1a barra mudando de sinal muda a serie que a
' carrega (carimbo novo: o refresh reaplica e a hachura acompanha); salvar e reabrir mantem a hachura e o estilo da pasta.
Public Sub Test_StyleApply_SurvivesValueSignChangeAndSaveReopen()
    Dim wb As Workbook, c As Chart, ws As Worksheet, p As String, stamp As String, s As Series, hatched As String, nm As String
    Set wb = NewTempWorkbook()
    Set c = NewWfChart(wb)
    Set ws = wb.Worksheets(1)
    Style_ApplyNamed c, "Cascata MC"
    stamp = Registry_GetOption(c, "stylestamp", "")
    ws.Range("B3").Value = 0.8
    Application.Calculate
    Core_SafeRefresh c
    AssertEqual stamp, Registry_GetOption(c, "stylestamp", ""), "valor mudou sem mudar a forma: carimbo igual"
    AssertEqual msoFillPatterned, c.SeriesCollection("tqq_uppos").Points(1).Format.Fill.Type, "hachura sobrevive a mudar valores"
    ws.Range("B2").Value = -20
    Application.Calculate
    Core_SafeRefresh c
    AssertTrue stamp <> Registry_GetOption(c, "stylestamp", ""), "1a barra negativa: outra serie a carrega"
    AssertEqual Style_Stamp(c), Registry_GetOption(c, "stylestamp", "")
    For Each s In c.SeriesCollection
        Select Case s.Name
            Case "tqq_uppos", "tqq_upneg", "tqq_downpos", "tqq_downneg", "tqq_total"
                If s.Points(1).Format.Fill.Type = msoFillPatterned Then hatched = hatched & s.Name & ","
        End Select
    Next
    AssertTrue Len(hatched) > 0 And InStr(hatched, "tqq_uppos,") = 0, "a hachura passou para a serie que carrega a barra: " & hatched
    AssertEqual msoFalse, c.SeriesCollection("tqq_uppos").Points(1).Format.Line.Visible, "a serie que perdeu a barra fica sem borda"
    nm = c.Parent.Name
    p = TempDir() & "\style-apply.xlsx"
    wb.SaveAs p, xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Set wb = Workbooks.Open(p)
    Set c = wb.Worksheets(1).ChartObjects(nm).Chart
    AssertEqual "Cascata MC", Registry_GetOption(c, "style", "")
    AssertEqual msoFillPatterned, c.SeriesCollection(Left$(hatched, InStr(hatched, ",") - 1)).Points(1).Format.Fill.Type, "hachura reabre"
    AssertTrue InStr(StyleStore_BookText(wb, "Cascata MC"), "primeira.padrao=diagonal_larga_para_cima") > 0, "o estilo reabre com a pasta"
    CloseTemp wb
    Kill p
End Sub

' Todos os 21 tipos aceitam um estilo de series sem erro; nos tipos de series, a 1a serie de dados recebe serie1.
Public Sub Test_StyleApply_AllKinds_NoError()
    Dim wb As Workbook, k As Variant, c As Chart, rng As Range, d As Object, ign As Long, s As Series, got As Boolean
    Set d = StyleDef_Parse(STYLEDEF_HEADER & vbCrLf & "serie1=#123456" & vbCrLf & "serie2=#654321" & vbCrLf & "fonte.tamanho=10", ign)
    For Each k In Array("stacked", "clustered", "stacked100", "line", "waterfall", "barstacked", "barclustered", "barstacked100", "barwaterfall", "area", "area100", "combo", "mekko", "mekkounits", "gantt", "pie", "doughnut", "butterfly", "scatter", "bubble", "waterfallpair")
        Set wb = NewTempWorkbook()
        Set rng = Samples_WriteTable(wb.Worksheets(1).Range("A1"), Samples_TableFor(CStr(k)))
        rng.Select
        Core_Insert CStr(k)
        Set c = Core_ActiveChart()
        AssertTrue Not c Is Nothing, CStr(k)
        Style_ApplyTo c, d
        AssertEqual 10, CLng(c.ChartArea.Font.Size), "fonte: " & CStr(k)
        got = False
        For Each s In c.SeriesCollection
            If s.Format.Fill.ForeColor.RGB = RGB(&H12, &H34, &H56) Or s.Format.Line.ForeColor.RGB = RGB(&H12, &H34, &H56) Then got = True
        Next
        If InStr(",stacked,clustered,stacked100,line,barstacked,barclustered,barstacked100,area,area100,combo,", "," & CStr(k) & ",") > 0 Then AssertTrue got, "serie1: " & CStr(k)
        CloseTemp wb
    Next
End Sub

Public Sub Test_StyleFromChart_CapturesManualColors()
    Dim wb As Workbook, c As Chart, t As String, d As Object, ign As Long
    Set wb = NewTempWorkbook()
    Set c = NewWfChart(wb)
    c.SeriesCollection("tqq_uppos").Format.Fill.ForeColor.RGB = RGB(1, 2, 3)        ' pintado a mao
    c.SeriesCollection("tqq_uppos").Points(1).Format.Fill.Patterned msoPatternDarkUpwardDiagonal
    t = Style_FromChart(c, "Do Grafico")
    Set d = StyleDef_Parse(t, ign)
    AssertEqual "#010203", StyleDef_Get(d, "aumento")
    AssertEqual "diagonal_escura_para_cima", StyleDef_Get(d, "primeira.padrao")
    AssertEqual "Do Grafico", StyleDef_Get(d, "nome")
    CloseTemp wb
End Sub

Public Sub Test_Core_StyleApply_Refusals_NeverRaise()
    Dim wb As Workbook, c As Chart
    Set wb = NewTempWorkbook()
    Set c = NewWfChart(wb)
    Core_ClearLastNotice
    Core_StyleApply "Nao Existe", "chart"
    AssertTrue InStr(Core_LastNotice(), "Estilo não encontrado") > 0, Core_LastNotice()
    Core_StyleApply "Cascata MC", "chart"
    AssertEqual "Cascata MC", Core_GetStyleName()
    Core_StyleApply "Tons de cinza", "chart"
    AssertEqual "Tons de cinza", Core_GetStyleName(), "paleta embutida vira o nome do estilo"
    CloseTemp wb
End Sub

Public Sub Test_StyleStore_Default_TestMemory()
    StyleStore_ResetTestHooks
    AssertEqual "", StyleStore_DefaultName()
    StyleStore_SetDefaultName "Cascata MC"
    AssertEqual "Cascata MC", StyleStore_DefaultName()
    StyleStore_ResetTestHooks
    AssertEqual "", StyleStore_DefaultName()
End Sub

' ---- Janela Estilos (1.12.0): textos e listas puras
Public Sub Test_StyleUi_RolesAndPicks()
    Dim k As Variant, items As Variant
    For Each k In StyleDef_Keys()
        AssertTrue Len(StyleUi_RoleLabel(CStr(k))) > 0, "rotulo em pt-BR: " & CStr(k)
        items = StyleUi_PickItems(CStr(k))
        If IsArray(items) Then
            AssertEqual StyleDef_Validate(CStr(k), StyleUi_PickValue(CStr(k), 0)), StyleUi_PickValue(CStr(k), 0), "1o item valido: " & CStr(k)
        End If
    Next
    AssertEqual "Aumentos (cascata)", StyleUi_RoleLabel("aumento")
    AssertEqual "#70AD47", StyleUi_PickValue("aumento", 9)
    AssertEqual "diagonal_larga_para_cima", StyleUi_PickValue("primeira.padrao", 3)
    AssertTrue StyleUi_IsColorKey("serie1")
    AssertFalse StyleUi_IsColorKey("espaco")
End Sub
