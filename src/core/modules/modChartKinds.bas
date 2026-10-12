Attribute VB_Name = "modChartKinds"
Option Explicit
' Fonte unica dos tipos de grafico e do que cada um suporta (faixa, Core_SetToggle/SetOption, Rebuild).
' A matriz e dado, nao logica: listas separadas por virgula consultadas com InStr (como WF_MARKERS).
Private Const KINDS As String = ",stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo,mekko,mekkounits,gantt,pie,doughnut,butterfly,scatter,bubble,"

' Os nomes sao procurados dentro de listas separadas por virgula: um nome com virgula ("stacked,clustered") casaria
' com dois tipos vizinhos da lista e seria aceito como se fosse um tipo. Por isso uma virgula recusa o nome.
Private Function InList(list As String, kind As String) As Boolean
    If Len(kind) = 0 Then Exit Function
    If InStr(kind, ",") > 0 Then Exit Function
    InList = InStr("," & list & ",", "," & kind & ",") > 0
End Function

Public Function Kinds_IsValid(kind As String) As Boolean
    If Len(kind) = 0 Then Exit Function
    If InStr(kind, ",") > 0 Then Exit Function
    Kinds_IsValid = InStr(KINDS, "," & kind & ",") > 0
End Function

' Barras horizontais: o eixo de categorias e vertical e o de valores horizontal.
Public Function Kinds_IsHorizontal(kind As String) As Boolean
    Kinds_IsHorizontal = InList("barstacked,barclustered,barstacked100,barwaterfall,gantt,butterfly", kind)
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
    ElseIf kind = "butterfly" Then
        Kinds_Family = "butterfly"
    ElseIf InList("scatter,bubble", kind) Then
        Kinds_Family = "scatter"
    End If
End Function

' Matriz de capacidades (design da Fase 3, secao 6). False para tipo ou recurso desconhecido. A matriz e estatica por tipo;
' a unica excecao e a legenda da cascata (waterfall/barwaterfall), que so existe com VARIAS series de dados: quem pergunta
' informa o numero de series (seriesCount; omitido = desconhecido = sem legenda). Core_Supports le o numero no bloco.
Public Function Kinds_Supports(kind As String, feature As String, Optional seriesCount As Long = 0) As Boolean
    Dim list As String, allKinds As String
    If Not Kinds_IsValid(kind) Then Exit Function
    allKinds = Mid$(KINDS, 2, Len(KINDS) - 2)
    Select Case feature
        Case "totals":     list = "stacked,stacked100,barstacked,barstacked100,area,area100,combo,mekko,mekkounits,butterfly"
        Case "labels":     list = allKinds
        Case "connectors": list = "stacked,stacked100,barstacked,barstacked100,waterfall,barwaterfall,combo"
        Case "numfmt":     list = "stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo,mekko,mekkounits,pie,doughnut,butterfly"
        Case "palette":    list = allKinds
        Case "gap":        list = "stacked,clustered,stacked100,barstacked,barclustered,barstacked100,waterfall,barwaterfall,combo,gantt,butterfly"
        Case "legend":     list = Replace(allKinds, ",gantt", "")      ' waterfall/barwaterfall: so com varias series (abaixo)
        Case "lineend":    list = "line,combo"                         ' legenda "nome no fim da linha" (Deco_SetLegend)
        Case "ind100":     list = "stacked100,barstacked100,area100,mekko"
        Case "vline":      list = "stacked,clustered,line,barstacked,barclustered,area,combo,mekkounits"
        Case "secondary":  list = "combo"
        Case "pieout":     list = "pie"                            ' a rosca nao aceita DataLabels.Position InsideEnd/OutsideEnd (1004)
        Case "cagr":       list = "stacked,clustered,stacked100,line,waterfall,barstacked,barclustered,barstacked100,barwaterfall,area,area100,combo"
        Case "diff":       list = "stacked,clustered,line,waterfall,barstacked,barclustered,barwaterfall,area,combo"
        Case "bracket":    list = "gantt"
        Case "samescale":  list = "stacked,clustered,line,waterfall,barstacked,barclustered,barwaterfall,area,combo,mekkounits,butterfly"
        Case "roundsum":   list = "stacked,stacked100,barstacked,barstacked100,area,area100"
        Case "break":      list = "stacked,clustered,barstacked,barclustered"
        Case "trend":      list = "scatter,bubble"
        Case "quad":       list = "scatter,bubble"
        Case Else: Exit Function
    End Select
    Kinds_Supports = InList(list, kind)
    If Kinds_Supports And feature = "legend" And Kinds_Family(kind) = "waterfall" Then Kinds_Supports = (seriesCount > 1)
End Function

' Rotulo (pt-BR) do controle da faixa que aciona o recurso: os avisos de "nao se aplica a este tipo de grafico" citam o botao que o
' usuario clicou. "" para recurso desconhecido.
Public Function Kinds_FeatureLabel(feature As String) As String
    Select Case feature
        Case "totals":     Kinds_FeatureLabel = "Totais"
        Case "labels":     Kinds_FeatureLabel = "Rótulos"
        Case "connectors": Kinds_FeatureLabel = "Conectores"
        Case "ind100":     Kinds_FeatureLabel = "Indicador 100%"
        Case "vline":      Kinds_FeatureLabel = "Linha de valor (média)"
        Case "secondary":  Kinds_FeatureLabel = "Eixo secundário (linha)"
        Case "pieout":     Kinds_FeatureLabel = "Rótulos fora da pizza"
        Case "gap":        Kinds_FeatureLabel = "Espaço"
        Case "legend":     Kinds_FeatureLabel = "Legenda"
        Case "numfmt":     Kinds_FeatureLabel = "Formato"
        Case "palette":    Kinds_FeatureLabel = "Paleta"
        Case "roundsum":   Kinds_FeatureLabel = "Rótulos com soma consistente"
        Case "break":      Kinds_FeatureLabel = "Corte de eixo"
        Case "trend":      Kinds_FeatureLabel = "Linha de tendência"
        Case "quad":       Kinds_FeatureLabel = "Quadrantes (médias)"
    End Select
End Function

' Texto unico do aviso "nao se aplica" (erro 7002) quando o tipo de grafico nao tem o recurso: cita o controle da faixa que o usuario clicou
' (Kinds_FeatureLabel); Espaco e Legenda tem frase propria. Usado pelos toggles e opcoes de modCoreApi, pela recusa de familia sem despacho, por
' Pie_SetOutside e por Deco_SetLegend (o mesmo texto em todo caminho que o usuario alcanca).
Public Function Kinds_NotApplicableText(feature As String) As String
    Select Case feature
        Case "gap": Kinds_NotApplicableText = "Espaço entre colunas não se aplica a este tipo de gráfico."
        Case "legend": Kinds_NotApplicableText = "Legenda não se aplica a este tipo de gráfico."
        Case Else: Kinds_NotApplicableText = "'" & Kinds_FeatureLabel(feature) & "' não se aplica a este tipo de gráfico."
    End Select
End Function

' O mesmo para um grafico concreto: le o tipo e, na cascata, o numero de series de dados do bloco. E a pergunta que a faixa
' (Core_Supports) e os comandos de opcao devem fazer: Kinds_Supports(tipo, recurso) sozinho nega a legenda da cascata.
Public Function Kinds_SupportsChart(cht As Chart, feature As String) As Boolean
    Dim kind As String, nSeries As Long
    kind = Registry_ChartType(cht)
    If Kinds_Family(kind) = "waterfall" Then nSeries = Waterfall_SeriesCount(cht)
    Kinds_SupportsChart = Kinds_Supports(kind, feature, nSeries)
End Function

' Numero de referencias de categoria que a primeira serie do grafico deve ter (Core_ChartSourceMismatch):
' o Mekko usa 5 pontos por categoria, os demais um.
Public Function Kinds_ExpectedCategoryRefs(kind As String, n As Long) As Long
    If Kinds_Family(kind) = "mekko" Then Kinds_ExpectedCategoryRefs = 5 * n Else Kinds_ExpectedCategoryRefs = n
End Function

' Despacho da insercao por familia. Cada Insert desfaz o que criou em caso de falha. Todas as familias validas tem
' implementacao (o Gantt foi a ultima); o ramo 7200 fica como protecao para um tipo valido sem Case (Core_Insert o
' trata como aviso ao usuario, sem registrar falha).
Public Function Kinds_Insert(ws As Worksheet, data As clsChartData, kind As String) As ChartObject
    Select Case Kinds_Family(kind)
        Case "stacked"
            Set Kinds_Insert = Stacked_Insert(ws, data, kind)
        Case "waterfall"
            Set Kinds_Insert = Waterfall_Insert(ws, data, kind)
        Case "area"
            Set Kinds_Insert = Area_Insert(ws, data, kind)
        Case "combo"
            Set Kinds_Insert = Combo_Insert(ws, data)
        Case "mekko"
            Set Kinds_Insert = Mekko_Insert(ws, data, kind)
        Case "pie"
            Set Kinds_Insert = Pie_Insert(ws, data, kind)
        Case "gantt"
            Set Kinds_Insert = Gantt_Insert(ws, data)
        Case "butterfly"
            Set Kinds_Insert = Bfly_Insert(ws, data)
        Case "scatter"
            Set Kinds_Insert = Scatter_Insert(ws, data, kind)
        Case Else
            Err.Raise vbObjectError + 7200, "Kinds_Insert", "Este tipo de gráfico ainda não está disponível nesta versão: " & kind
    End Select
End Function
