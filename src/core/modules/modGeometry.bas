Attribute VB_Name = "modGeometry"
Option Explicit
' Converte categoria/valor em coordenadas (pontos) dentro do gráfico, a partir da área de plotagem
' interna e da escala do eixo de valores. Funções puras primeiro (testáveis sem gráfico).

Public Function Geom_SlotWidth(insideWidth As Double, nCats As Long) As Double
    Geom_SlotWidth = insideWidth / nCats
End Function

Public Function Geom_CategoryCenterX(insideLeft As Double, insideWidth As Double, nCats As Long, idx As Long, Optional between As Boolean = True) As Double
    If between Then
        Geom_CategoryCenterX = insideLeft + (idx - 0.5) * insideWidth / nCats
    ElseIf nCats <= 1 Then
        Geom_CategoryCenterX = insideLeft + insideWidth / 2
    Else
        Geom_CategoryCenterX = insideLeft + (idx - 1) * insideWidth / (nCats - 1)
    End If
End Function

Public Function Geom_BarWidth(slotWidth As Double, gapWidthPct As Double) As Double
    Geom_BarWidth = slotWidth / (1 + gapWidthPct / 100)
End Function

Public Function Geom_ValueToY(insideTop As Double, insideHeight As Double, axisMin As Double, axisMax As Double, v As Double) As Double
    If axisMax = axisMin Then
        Geom_ValueToY = insideTop + insideHeight
    Else
        Geom_ValueToY = insideTop + insideHeight * (axisMax - v) / (axisMax - axisMin)
    End If
End Function

' Períodos entre duas categorias: anos se forem datas ou anos (1900..2200); senão diferença de índices.
Public Function Geom_Periods(catFrom As Variant, catTo As Variant, idxFrom As Long, idxTo As Long) As Double
    Dim p As Double
    If VarType(catFrom) = vbDate And VarType(catTo) = vbDate Then
        p = (CDate(catTo) - CDate(catFrom)) / 365.25
    ElseIf IsNumeric(catFrom) And IsNumeric(catTo) And Not IsEmpty(catFrom) And Not IsEmpty(catTo) Then
        If CDbl(catFrom) >= 1900 And CDbl(catFrom) <= 2200 And CDbl(catTo) >= 1900 And CDbl(catTo) <= 2200 Then
            p = CDbl(catTo) - CDbl(catFrom)
        End If
    End If
    If p <= 0 Then p = idxTo - idxFrom
    Geom_Periods = p
End Function

Public Function Geom_ChartCategoryCount(cht As Chart) As Long
    If cht.SeriesCollection.Count = 0 Then Exit Function
    Geom_ChartCategoryCount = cht.SeriesCollection(1).Points.Count
End Function

Public Function Geom_ChartCategoryCenterX(cht As Chart, idx As Long) As Double
    With cht.PlotArea
        Geom_ChartCategoryCenterX = Geom_CategoryCenterX(.InsideLeft, .InsideWidth, Geom_ChartCategoryCount(cht), idx, Geom_ChartBetween(cht))
    End With
End Function


Public Function Geom_ChartValueToY(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue)
    With cht.PlotArea
        Geom_ChartValueToY = Geom_ValueToY(.InsideTop, .InsideHeight, ax.MinimumScale, ax.MaximumScale, v)
    End With
End Function

Public Function Geom_ChartBarHalfWidth(cht As Chart) As Double
    Dim gap As Double
    If Registry_ChartType(cht) = "line" Then Exit Function    ' linha: sem barras, ancora no ponto
    gap = 50
    On Error Resume Next
    gap = cht.ChartGroups(1).GapWidth
    On Error GoTo 0
    Geom_ChartBarHalfWidth = Geom_BarWidth(Geom_SlotWidth(cht.PlotArea.InsideWidth, Geom_ChartCategoryCount(cht)), gap) / 2
End Function

' True quando os pontos ficam no meio das fatias (padrão de colunas); False quando o eixo cruza nas categorias
' (linha com AxisBetweenCategories = False: primeiro ponto na borda esquerda, último na direita).
Public Function Geom_ChartBetween(cht As Chart) As Boolean
    Geom_ChartBetween = True
    On Error Resume Next
    Geom_ChartBetween = cht.Axes(xlCategory).AxisBetweenCategories
    On Error GoTo 0
End Function
