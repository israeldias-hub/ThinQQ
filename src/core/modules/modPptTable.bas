Attribute VB_Name = "modPptTable"
Option Explicit
' Tabela nativa do PowerPoint vinculada a um intervalo do Excel (1.10.0; "Create tables with Excel formatting" do think-cell; spec 2026-10-15
' secao 5.4). A ultima matriz de textos escrita fica na etiqueta TQQ_TBLAST: a atualizacao escreve so as celulas cuja ORIGEM mudou (a
' formatacao feita no PowerPoint fica). Limites do PowerPoint: 75 linhas e 75 colunas (o ThinQQ limita a 25 colunas).
' ETIQUETA (adendo do controlador, medido na 1.6.0): o .pptx salvo perde Chr(30)/Chr(31) e outros controles, troca Tab por espaco e CR por LF;
' LF e "|" sobrevivem. TQQ_TBLAST = linhas separadas por LF, celulas por "|", cada celula codificada por PptText_TagEnc ("\\" "\r" "\n" "\t"
' "\xHH") e com "|" -> "\x7C": so caracteres imprimiveis e os separadores, ida e volta sem perda.
Public Const PPTTABLE_TAG As String = "TQQ_TABLE"
Public Const PPTTABLE_TAG_LAST As String = "TQQ_TBLAST"
Public Const PPTTABLE_MAX_ROWS As Long = 75
Public Const PPTTABLE_MAX_COLS As Long = 25
Public Const PPTTABLE_MAX_TAG As Long = 200000               ' TQQ_TBLAST maior que isto nao e gravada (a proxima atualizacao reescreve tudo)
Private Const PPTTABLE_RS As String = vbLf                    ' separador de linhas na etiqueta
Private Const PPTTABLE_FS As String = "|"                     ' separador de celulas na etiqueta

' Matriz 1..r x 1..c de textos -> etiqueta (ver cabecalho). Pura.
Public Function PptTable_Serialize(m As Variant) As String
    Dim r As Long, c As Long, rowsOut() As String, cells() As String
    ReDim rowsOut(1 To UBound(m, 1))
    ReDim cells(1 To UBound(m, 2))
    For r = 1 To UBound(m, 1)
        For c = 1 To UBound(m, 2)
            cells(c) = Replace(PptText_TagEnc(CStr(m(r, c))), PPTTABLE_FS, "\x7C")
        Next
        rowsOut(r) = Join(cells, PPTTABLE_FS)
    Next
    PptTable_Serialize = Join(rowsOut, PPTTABLE_RS)
End Function

' Inverso de PptTable_Serialize: matriz 1..r x 1..c de String (colunas = as da 1a linha; linha curta = celulas vazias) ou Empty (etiqueta
' vazia, maior que PPTTABLE_MAX_TAG ou com mais de PPTTABLE_MAX_ROWS linhas / PPTTABLE_MAX_COLS colunas). A etiqueta vem da apresentacao (de
' quem a fez): nunca levanta erro. Pura.
Public Function PptTable_Parse(s As String) As Variant
    Dim rows() As String, cells() As String, m() As String, r As Long, c As Long, nc As Long
    On Error GoTo Bad
    If Len(s) = 0 Or Len(s) > PPTTABLE_MAX_TAG Then Exit Function
    rows = Split(s, PPTTABLE_RS)
    If UBound(rows) + 1 > PPTTABLE_MAX_ROWS Then Exit Function
    nc = UBound(Split(rows(0), PPTTABLE_FS)) + 1
    If nc > PPTTABLE_MAX_COLS Then Exit Function
    ReDim m(1 To UBound(rows) + 1, 1 To nc)
    For r = 0 To UBound(rows)
        cells = Split(rows(r), PPTTABLE_FS)
        For c = 0 To nc - 1
            If c <= UBound(cells) Then m(r + 1, c + 1) = PptText_TagDec(cells(c))
        Next
    Next
    PptTable_Parse = m
    Exit Function
Bad:
    PptTable_Parse = Empty
End Function

' Celulas de newM a escrever: diferentes de oldM ou fora das dimensoes dele (oldM Empty = todas). Itens Array(linha, coluna). Pura.
Public Function PptTable_Diff(oldM As Variant, newM As Variant) As Collection
    Dim out As New Collection, r As Long, c As Long, inOld As Boolean
    For r = 1 To UBound(newM, 1)
        For c = 1 To UBound(newM, 2)
            inOld = False
            If Not IsEmpty(oldM) Then inOld = (r <= UBound(oldM, 1) And c <= UBound(oldM, 2))
            If Not inOld Then
                out.Add Array(r, c)
            ElseIf CStr(oldM(r, c)) <> CStr(newM(r, c)) Then
                out.Add Array(r, c)
            End If
        Next
    Next
    Set PptTable_Diff = out
End Function

' Texto seguro para uma celula da tabela: quebra de linha da celula (CR, LF, CRLF: Alt+Enter) vira a quebra de linha do PowerPoint (Chr(11));
' outros caracteres de controle viram espaco; cortado em PPTTEXT_MAX_LEN sem partir par substituto (emoji). Pura.
Public Function PptTable_SafeCell(t As String) As String
    Dim s As String, i As Long, c As Long
    s = Replace(Replace(Replace(t, vbCrLf, vbLf), vbCr, vbLf), vbLf, Chr$(11))
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If ((c >= 0 And c < 32) Or c = 127) And c <> 11 Then Mid$(s, i, 1) = " "
    Next
    If Len(s) > PPTTEXT_MAX_LEN Then
        c = AscW(Mid$(s, PPTTEXT_MAX_LEN, 1)) And &HFFFF&
        If c >= &HD800& And c <= &HDBFF& Then s = Left$(s, PPTTEXT_MAX_LEN - 1) Else s = Left$(s, PPTTEXT_MAX_LEN)
    End If
    PptTable_SafeCell = s
End Function

' Textos exibidos das celulas de rng (PptText_CellText: formato do Excel, coluna estreita sem "####"), saneados (PptTable_SafeCell). So le.
Public Function PptTable_TextMatrix(rng As Range) As Variant
    Dim m() As String, r As Long, c As Long
    ReDim m(1 To rng.Rows.Count, 1 To rng.Columns.Count)
    For r = 1 To rng.Rows.Count
        For c = 1 To rng.Columns.Count
            m(r, c) = PptTable_SafeCell(PptText_CellText(rng.Cells(r, c)))
        Next
    Next
    PptTable_TextMatrix = m
End Function

' "" = o intervalo serve; senao o motivo (varias areas, tamanho, mescladas).
Public Function PptTable_RefusalText(rng As Range) As String
    If rng.Areas.Count > 1 Then PptTable_RefusalText = "Selecione um único intervalo retangular.": Exit Function
    If rng.Rows.Count > PPTTABLE_MAX_ROWS Or rng.Columns.Count > PPTTABLE_MAX_COLS Then PptTable_RefusalText = "A tabela vinculada aceita até " & PPTTABLE_MAX_ROWS & " linhas e " & PPTTABLE_MAX_COLS & " colunas.": Exit Function
    If IsNull(rng.MergeCells) Then
        PptTable_RefusalText = "A tabela vinculada não aceita células mescladas no intervalo."
    ElseIf rng.MergeCells Then
        PptTable_RefusalText = "A tabela vinculada não aceita células mescladas no intervalo."
    End If
End Function
