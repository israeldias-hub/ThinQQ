VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Tabela de entrada: linha 1 = cabeçalho (canto + nomes das séries); coluna 1 = categorias.
Private mSource As Range

Public Sub Init(src As Range)
    Set mSource = src
End Sub

Public Property Get Source() As Range
    Set Source = mSource
End Property

Public Property Get CategoryCount() As Long
    CategoryCount = mSource.Rows.Count - 1
End Property

Public Property Get SeriesCount() As Long
    SeriesCount = mSource.Columns.Count - 1
End Property

Public Function CategoryCells() As Range
    Set CategoryCells = mSource.Offset(1, 0).Resize(CategoryCount, 1)
End Function

Public Function SeriesNameCell(j As Long) As Range
    Set SeriesNameCell = mSource.Cells(1, j + 1)
End Function

Public Function SeriesValueCells(j As Long) As Range
    Set SeriesValueCells = mSource.Offset(1, j).Resize(CategoryCount, 1)
End Function

Public Function ValueCell(i As Long, j As Long) As Range
    Set ValueCell = mSource.Cells(i + 1, j + 1)
End Function

Public Function CategoryText(i As Long) As String
    CategoryText = CStr(mSource.Cells(i + 1, 1).Text)
End Function

Public Function SeriesName(j As Long) As String
    SeriesName = CStr(SeriesNameCell(j).Text)
End Function

Public Function Validate() As String
    If mSource Is Nothing Then Validate = "Selecione a tabela de dados.": Exit Function
    If mSource.Areas.Count > 1 Then Validate = "Selecione uma única área contígua.": Exit Function
    If mSource.Rows.Count < 2 Or mSource.Columns.Count < 2 Then
        Validate = "Selecione a tabela com a linha de cabeçalho (nomes das séries) e a coluna de categorias (mínimo 2 linhas x 2 colunas)."
        Exit Function
    End If
    If mSource.Worksheet.Name = CALC_SHEET Then Validate = "Essa é a planilha interna do ThinQQ.": Exit Function
    Validate = ""
End Function
