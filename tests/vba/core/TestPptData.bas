Attribute VB_Name = "TestPptData"
Option Explicit
' Parte G (1.8.0): pasta de dados dos graficos criados pelo PowerPoint (sem PowerPoint).

Private Function SheetExists(wb As Workbook, nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = wb.Sheets(nm)
    SheetExists = Not s Is Nothing
End Function

Public Sub Test_PptData_PathFor()
    AssertEqual "C:\a\Vendas - dados ThinQQ.xlsx", PptData_PathFor("C:\a\Vendas.pptx")
    AssertEqual "C:\a\b.c - dados ThinQQ.xlsx", PptData_PathFor("C:\a\b.c.pptx"), "so a ultima extensao sai"
    AssertEqual "\\srv\x\R - dados ThinQQ.xlsx", PptData_PathFor("\\srv\x\R.pptx")
    AssertEqual "", PptData_PathFor("https://empresa.sharepoint.com/x/R.pptx"), "endereco da Web: recusado"
    AssertEqual "", PptData_PathFor("Apresentação1"), "nunca salva: sem pasta"
End Sub

Public Sub Test_PptData_NextSheetName()
    AssertEqual "Gráfico 1", PptData_NextSheetName("|Leia-me|")
    AssertEqual "Gráfico 3", PptData_NextSheetName("|Leia-me|Gráfico 1|Gráfico 2|")
    AssertEqual "Gráfico 2", PptData_NextSheetName("|Gráfico 1|Gráfico 3|"), "primeiro livre"
End Sub

' Tipo desconhecido e recusado antes de qualquer acesso ao slide (nem PowerPoint e preciso); nunca propaga erro.
Public Sub Test_PptInsertChart_UnknownKind_NoPowerPoint()
    AssertEqual "Tipo de gráfico desconhecido: xyz", Core_PptInsertChart(Nothing, "xyz")
End Sub

Public Sub Test_PptData_CreateReopenAndNewChart()
    Dim p As String, wb As Workbook, co As ChartObject, rng As Range
    p = TempDir() & "\Teste G - dados ThinQQ.xlsx"
    If Len(Dir(p)) > 0 Then Kill p
    Set wb = PptData_OpenOrCreate(p, "Teste G.pptx")
    AssertTrue PptData_IsDataBook(wb), "marcada"
    AssertEqual "Leia-me", wb.Worksheets(1).Name
    AssertTrue Len(Dir(p)) > 0, "salva no caminho"
    Set co = PptData_NewChart(wb, "stacked", rng)
    AssertEqual "Gráfico 1", co.Parent.Name
    AssertEqual "$A$1:$D$5", rng.Address, "tabela de exemplo (Ano x 3 regioes, 4 anos)"
    AssertEqual "stacked", Registry_ChartType(co.Chart)
    Set co = PptData_NewChart(wb, "pie", rng)
    AssertEqual "Gráfico 2", co.Parent.Name
    wb.Save
    AssertTrue PptData_OpenOrCreate(p, "Teste G.pptx") Is wb, "ja aberta: a mesma"
    Watchers_Stop
    wb.Close SaveChanges:=False
    Set wb = PptData_OpenForEdit(p)
    AssertFalse wb.ReadOnly
    AssertTrue SheetExists(wb, "Gráfico 1") And SheetExists(wb, "Gráfico 2"), "reaberta com os dois graficos"
    wb.Close SaveChanges:=False
    Kill p
End Sub
