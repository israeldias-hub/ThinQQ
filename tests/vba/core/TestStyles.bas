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
    p = TempDir() & "\entrada.tqqestilo"                       ' fora do deposito (senao o proprio arquivo ja seria "o estilo de mesmo nome")
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

Private Function MarkerBg(s As Series) As Long
    MarkerBg = -1
    On Error Resume Next
    MarkerBg = s.MarkerBackgroundColor          ' bolhas e auxiliares sem marcador recusam a propriedade
    Err.Clear
End Function

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
            If Kinds_Family(CStr(k)) = "pie" Then
                If s.Points(1).Format.Fill.ForeColor.RGB = RGB(&H12, &H34, &H56) Then got = True        ' fatia 1
            ElseIf Kinds_Family(CStr(k)) = "scatter" Then
                If MarkerBg(s) = RGB(&H12, &H34, &H56) Then got = True                                  ' marcador do grupo
            End If
        Next
        ' revisao final 1.12.0 (M11): a cor confere em todo tipo com serieN (fora: Gantt, que so recebe fonte/rotulos/espaco, e a cascata de
        ' uma serie, que usa aumento/reducao/total)
        If InStr(",gantt,waterfall,barwaterfall,", "," & CStr(k) & ",") = 0 Then AssertTrue got, "serie1: " & CStr(k)
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

' ---- Revisao final da 1.12.0 (onda de correcao)

Private Function NewStackedChart(wb As Workbook) As Chart
    Dim rng As Range
    Set rng = Samples_WriteTable(wb.Worksheets(1).Range("A1"), Samples_TableFor("stacked"))
    rng.Select
    Core_Insert "stacked"
    Set NewStackedChart = Core_ActiveChart()
End Function

Private Sub WriteStyleFile(p As String, lines As Variant)
    Dim f As Integer, i As Long
    f = FreeFile
    Open p For Output As #f
    For i = LBound(lines) To UBound(lines)
        Print #f, CStr(lines(i))
    Next
    Close #f
End Sub

' I1: importar um arquivo com o nome de um estilo seu NAO grava sem perguntar ("?" + texto); com allowOverwrite substitui.
Public Sub Test_StyleStore_Import_SameName_AsksBeforeReplacing()
    Dim d As String, p As String, nm As String, ign As Long, r As String
    StyleStore_ResetTestHooks
    d = StyleTestDir()
    AssertEqual "", StyleStore_SaveUser("Relatorio Mensal", STYLEDEF_HEADER & vbCrLf & "aumento=#111111")
    p = TempDir() & "\relatorio.tqqestilo"
    WriteStyleFile p, Array(STYLEDEF_HEADER, "nome=Relatorio Mensal", "aumento=#222222")
    r = StyleStore_Import(p, nm, ign)
    AssertEqual "?", Left$(r, 1), "mesmo nome: pergunta (nada gravado)"
    AssertEqual "Relatorio Mensal", nm
    AssertTrue InStr(Mid$(r, 2), "aumento=#222222") > 0, "o texto canonico do arquivo volta para quem pergunta"
    AssertTrue InStr(StyleStore_UserText("Relatorio Mensal"), "#111111") > 0, "o estilo do usuario ficou intacto"
    AssertTrue InStr(StyleStore_ReplaceQuestion(nm), "Você já tem um estilo chamado 'Relatorio Mensal'. Substituir pelo do arquivo?") > 0
    r = StyleStore_Import(p, nm, ign, True)
    AssertTrue Left$(r, 1) <> "?" And Left$(r, 1) <> "!", r
    AssertTrue InStr(StyleStore_UserText("Relatorio Mensal"), "#222222") > 0, "com allowOverwrite: substitui"
    Kill p
    StyleStore_ResetTestHooks
End Sub

' I2 e M1: pasta hostil. Uma parte com 150 estilos, uma malformada, um <style> sem nome e um "Cascata MC" de outras cores: le no maximo 100,
' a lista da faixa tem no maximo 200 nomes, o embutido nunca e sombreado e gravar mais um estilo na pasta cheia e recusado (sem travar).
Public Sub Test_StyleStore_HostileBook_CapsAndBuiltinsWin()
    Dim wb As Workbook, parts() As String, i As Long, t As Single, c As Chart, k As Long, en As Long
    Set wb = NewTempWorkbook()
    ReDim parts(0 To 151)
    parts(0) = "<tqqStyles xmlns=""" & STYLESTORE_NS & """><style>sem nome</style><style name=""Cascata MC"">" & STYLEDEF_HEADER & "&#10;aumento=#000001</style>"
    For i = 1 To 150
        parts(i) = "<style name=""s" & i & """>" & STYLEDEF_HEADER & "</style>"
    Next
    parts(151) = "</tqqStyles>"
    wb.CustomXMLParts.Add Join(parts, "")
    On Error Resume Next
    wb.CustomXMLParts.Add "<tqqStyles xmlns=""" & STYLESTORE_NS & """><style name=""x"">"         ' malformada: o Office recusa ou fica sem filhos
    Err.Clear
    On Error GoTo 0
    t = Timer
    AssertEqual STYLESTORE_BOOK_MAX, StyleStore_BookAll(wb).Count, "no maximo 100 estilos lidos da pasta"
    AssertTrue UBound(Split(StyleStore_Names(wb), "|")) + 1 <= STYLESTORE_NAMES_MAX, "no maximo 200 nomes para a faixa"
    AssertTrue InStr(StyleStore_Get("Cascata MC", wb, False), "aumento=#70AD47") > 0, "o embutido vence a copia da pasta (M1)"
    On Error Resume Next
    StyleStore_SaveBook wb, "Mais um", STYLEDEF_HEADER
    en = Err.Number
    On Error GoTo 0
    AssertEqual vbObjectError + 7140, en, "pasta cheia: gravar mais um e recusado"
    ' grafico com style forjado (estilo inexistente) e carimbo forjado: cada refresh tenta reaplicar e falha rapido, sem travar
    Set c = NewStackedChart(wb)
    Registry_SetOption c, "style", "Nao Existe"
    Registry_SetOption c, "stylestamp", "forjado"
    For k = 1 To 10
        Core_SafeRefresh c
    Next
    AssertTrue Timer - t < 20, "pasta hostil: leitura, lista e 10 refresh em menos de 20 s (" & Format$(Timer - t, "0.0") & " s)"
    CloseTemp wb
End Sub

' M1/M10d: nome de dispositivo do Windows e arquivo posto a mao com nome invalido ou de embutido nao entram.
Public Sub Test_StyleStore_DeviceAndForeignNames()
    Dim d As String
    StyleStore_ResetTestHooks
    d = StyleTestDir()
    AssertTrue StyleStore_IsDeviceName("con") And StyleStore_IsDeviceName("COM1") And StyleStore_IsDeviceName("lpt9")
    AssertFalse StyleStore_IsDeviceName("CONTA") Or StyleStore_IsDeviceName("COMX")
    AssertTrue InStr(StyleStore_SaveUser("NUL", STYLEDEF_HEADER), "nome reservado do Windows") > 0
    WriteStyleFile d & "\a;b" & STYLEDEF_EXT, Array(STYLEDEF_HEADER)
    WriteStyleFile d & "\Cascata MC" & STYLEDEF_EXT, Array(STYLEDEF_HEADER, "aumento=#000001")
    AssertEqual 0, InStr("|" & StyleStore_Names() & "|", "|a;b|"), "arquivo com ; fora da lista"
    AssertEqual 1, UBound(Split("|" & StyleStore_Names() & "|", "|Cascata MC|")), "Cascata MC uma vez so (o embutido)"
    AssertTrue InStr(StyleStore_Get("Cascata MC", Nothing, True), "aumento=#70AD47") > 0, "o arquivo posto a mao nao sombreia o embutido"
    Kill d & "\a;b" & STYLEDEF_EXT
    Kill d & "\Cascata MC" & STYLEDEF_EXT
    StyleStore_ResetTestHooks
End Sub

' I3: tirar o estilo (paleta embutida) marca stylever = "pal": o slide continua recebendo a formatacao (da paleta); a reaplicacao nao age.
Public Sub Test_StyleClear_MarksPaletteVersion()
    Dim wb As Workbook, c As Chart
    Set wb = NewTempWorkbook()
    Set c = NewWfChart(wb)
    Style_ApplyNamed c, "Cascata MC"
    Style_ApplyNamed c, "Tons de cinza"
    AssertEqual "", Registry_GetOption(c, "style", "")
    AssertEqual STYLE_VER_PALETTE, Registry_GetOption(c, "stylever", "")
    AssertEqual "Tons de cinza", Core_GetStyleName()
    Style_ReapplyIfNeeded c
    AssertEqual "", Registry_GetOption(c, "style", ""), "sem style: nada reaplicado"
    Style_ApplyNamed c, "Cascata DDE"
    AssertEqual 8, Len(Registry_GetOption(c, "stylever", "")), "estilo de novo: a versao do texto"
    CloseTemp wb
End Sub

' I5: o espaco do estilo vale no aplicar explicito; a lista Espaco continua mudando o grafico depois (a reaplicacao nao o desfaz).
Public Sub Test_StyleApply_Gap_OnlyOnExplicitApply()
    Dim wb As Workbook, c As Chart
    Set wb = NewTempWorkbook()
    Set c = NewStackedChart(wb)
    StyleStore_SaveBook wb, "Com Espaco", STYLEDEF_HEADER & vbCrLf & "nome=Com Espaco" & vbCrLf & "serie1=#123456" & vbCrLf & "espaco=50"
    Style_ApplyNamed c, "Com Espaco"
    AssertEqual 50, c.ChartGroups(1).GapWidth, "aplicar: o espaco do estilo"
    Core_SetOption "gap", "120"
    AssertEqual 120, c.ChartGroups(1).GapWidth, "Espaco > 120 depois do estilo"
    AssertEqual "120", Registry_GetOption(c, "gap", "")
    Core_SetToggle "labels", False
    AssertEqual 120, c.ChartGroups(1).GapWidth, "a reaplicacao (alternador) nao volta ao 50"
    AssertEqual RGB(&H12, &H34, &H56), CLng(c.SeriesCollection(1).Format.Fill.ForeColor.RGB), "as cores continuam reaplicadas"
    CloseTemp wb
End Sub

' M5: rotulo religado pelo refresh (9 pt) volta ao rotulo.tamanho do estilo sem o carimbo mudar.
Public Sub Test_StyleApply_RelinkedLabels_GetStyleSize()
    Dim wb As Workbook, c As Chart, ws As Worksheet
    Set wb = NewTempWorkbook()
    Set c = NewStackedChart(wb)
    Set ws = wb.Worksheets(1)
    StyleStore_SaveBook wb, "Rotulo 12", STYLEDEF_HEADER & vbCrLf & "nome=Rotulo 12" & vbCrLf & "rotulo.tamanho=12"
    Style_ApplyNamed c, "Rotulo 12"
    AssertEqual 12, CLng(c.SeriesCollection(1).Points(1).DataLabel.Font.Size), "aplicado"
    Labels_TakeRelinked
    Application.EnableEvents = False                           ' sem o refresh do observador no meio (ele ja refaria o estilo)
    Labels_LinkPoints c.SeriesCollection(1), ws.Range("B2:B5"), "0", xlLabelPositionCenter
    Application.EnableEvents = True
    AssertEqual 9, CLng(c.SeriesCollection(1).Points(1).DataLabel.Font.Size), "religado com 9 pt"
    Core_SafeRefresh c
    AssertEqual 12, CLng(c.SeriesCollection(1).Points(1).DataLabel.Font.Size), "o refresh refaz rotulo.tamanho"
    CloseTemp wb
End Sub

' M10a: Novo do grafico numa pizza le a cor de cada fatia.
Public Sub Test_StyleFromChart_Pie_ReadsSlices()
    Dim wb As Workbook, c As Chart, rng As Range, d As Object, ign As Long
    Set wb = NewTempWorkbook()
    Set rng = Samples_WriteTable(wb.Worksheets(1).Range("A1"), Samples_TableFor("pie"))
    rng.Select
    Core_Insert "pie"
    Set c = Core_ActiveChart()
    Style_ApplyNamed c, "Cascata MC"
    Set d = StyleDef_Parse(Style_FromChart(c, "Da Pizza"), ign)
    AssertEqual "#70AD47", StyleDef_Get(d, "serie1")
    AssertEqual "#D9D9D9", StyleDef_Get(d, "serie2")
    AssertEqual "#E8173C", StyleDef_Get(d, "serie3")
    CloseTemp wb
End Sub
