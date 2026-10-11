Attribute VB_Name = "modChartKinds"
Option Explicit
' Fonte unica dos tipos de grafico e do que cada um suporta (faixa, Core_SetToggle/SetOption, Rebuild).
' A matriz e dado, nao logica: listas separadas por virgula consultadas com InStr (como WF_MARKERS).
Private Const KINDS As String = ",stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo,mekko,mekkounits,gantt,pie,doughnut,"

Private Function InList(list As String, kind As String) As Boolean
    If Len(kind) = 0 Then Exit Function
    InList = InStr("," & list & ",", "," & kind & ",") > 0
End Function

Public Function Kinds_IsValid(kind As String) As Boolean
    If Len(kind) = 0 Then Exit Function
    Kinds_IsValid = InStr(KINDS, "," & kind & ",") > 0
End Function

' Barras horizontais: o eixo de categorias e vertical e o de valores horizontal.
Public Function Kinds_IsHorizontal(kind As String) As Boolean
    Kinds_IsHorizontal = InList("barstacked,barclustered,barstacked100,barwaterfall,gantt", kind)
End Function

' Tipos cujos segmentos mostram participacao (valor / total da linha).
Public Function Kinds_Is100(kind As String) As Boolean
    Kinds_Is100 = InList("stacked100,barstacked100,area100,mekko", kind)
End Function

' Familia de construcao (modulo que implementa o Insert); "" se o tipo e invalido.
Public Function Kinds_Family(kind As String) As String
    If InList("stacked,clustered,stacked100,line,barstacked,barclustered,barstacked100", kind) Then
        Kinds_Family = "stacked"
    ElseIf InList("waterfall,barwaterfall", kind) Then
        Kinds_Family = "waterfall"
    ElseIf InList("area,area100", kind) Then
        Kinds_Family = "area"
    ElseIf kind = "combo" Then
        Kinds_Family = "combo"
    ElseIf InList("mekko,mekkounits", kind) Then
        Kinds_Family = "mekko"
    ElseIf kind = "gantt" Then
        Kinds_Family = "gantt"
    ElseIf InList("pie,doughnut", kind) Then
        Kinds_Family = "pie"
    End If
End Function

' Matriz de capacidades (design da Fase 2, secao 7). False para tipo ou recurso desconhecido.
Public Function Kinds_Supports(kind As String, feature As String) As Boolean
    Dim list As String, allKinds As String
    If Not Kinds_IsValid(kind) Then Exit Function
    allKinds = Mid$(KINDS, 2, Len(KINDS) - 2)
    Select Case feature
        Case "totals":     list = "stacked,stacked100,barstacked,barstacked100,area,area100,combo,mekko,mekkounits"
        Case "labels":     list = allKinds
        Case "connectors": list = "stacked,stacked100,barstacked,barstacked100,waterfall,barwaterfall,combo"
        Case "numfmt":     list = Replace(allKinds, ",gantt", "")
        Case "palette":    list = allKinds
        Case "gap":        list = "stacked,clustered,stacked100,barstacked,barclustered,barstacked100,waterfall,barwaterfall,combo,gantt"
        Case "legend":     list = Replace(allKinds, ",gantt", "")      ' em waterfall/barwaterfall so com varias series (a Tarefa 4 decide)
        Case "ind100":     list = "stacked100,barstacked100,area100,mekko"
        Case "vline":      list = "stacked,clustered,line,barstacked,barclustered,area,combo,mekkounits"
        Case "secondary":  list = "combo"
        Case "pieout":     list = "pie"                            ' a rosca nao aceita DataLabels.Position InsideEnd/OutsideEnd (1004)
        Case "cagr":       list = "stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo"
        Case "diff":       list = "stacked,clustered,line,waterfall,barstacked,barclustered,barwaterfall,area,combo"
        Case "bracket":    list = "gantt"
        Case "samescale":  list = "stacked,clustered,line,waterfall,barstacked,barclustered,barwaterfall,area,combo,mekkounits"
        Case Else: Exit Function
    End Select
    Kinds_Supports = InList(list, kind)
End Function

' Numero de referencias de categoria que a primeira serie do grafico deve ter (Core_ChartSourceMismatch):
' o Mekko usa 5 pontos por categoria, os demais um.
Public Function Kinds_ExpectedCategoryRefs(kind As String, n As Long) As Long
    If Kinds_Family(kind) = "mekko" Then Kinds_ExpectedCategoryRefs = 5 * n Else Kinds_ExpectedCategoryRefs = n
End Function

' Despacho da insercao por familia. Cada Insert desfaz o que criou em caso de falha. As tarefas seguintes
' acrescentam os Case das demais familias; ate la, o tipo valido sem implementacao avisa com o erro 7200
' (tratado por Core_Insert como aviso ao usuario, sem registrar falha).
Public Function Kinds_Insert(ws As Worksheet, data As clsChartData, kind As String) As ChartObject
    Select Case Kinds_Family(kind)
        Case "stacked"
            ' Barras horizontais so existem a partir da Tarefa 3 (esta guarda sai la): sem ela o Stacked_Insert
            ' criaria colunas com o tipo gravado como "barstacked".
            If Kinds_IsHorizontal(kind) Then
                Err.Raise vbObjectError + 7200, "Kinds_Insert", "Este tipo de gráfico ainda não está disponível nesta versão: " & kind
            End If
            Set Kinds_Insert = Stacked_Insert(ws, data, kind)
        Case "waterfall"
            If kind <> "waterfall" Then
                Err.Raise vbObjectError + 7200, "Kinds_Insert", "Este tipo de gráfico ainda não está disponível nesta versão: " & kind
            End If
            Set Kinds_Insert = Waterfall_Insert(ws, data)
        Case Else
            Err.Raise vbObjectError + 7200, "Kinds_Insert", "Este tipo de gráfico ainda não está disponível nesta versão: " & kind
    End Select
End Function
