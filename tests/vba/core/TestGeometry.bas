Attribute VB_Name = "TestGeometry"
Option Explicit

Public Sub Test_Geom_CategoryCenterX()
    AssertAlmost 150, Geom_CategoryCenterX(100, 400, 4, 1)   ' slot de 100: centro em 100+50
    AssertAlmost 450, Geom_CategoryCenterX(100, 400, 4, 4)
End Sub

Public Sub Test_Geom_BarWidth()
    AssertAlmost 100, Geom_SlotWidth(400, 4)
    AssertAlmost 66.6666667, Geom_BarWidth(100, 50)
    AssertAlmost 100, Geom_BarWidth(100, 0)
End Sub

Public Sub Test_Geom_ValueToY()
    AssertAlmost 50, Geom_ValueToY(50, 200, 0, 100, 100)        ' topo
    AssertAlmost 250, Geom_ValueToY(50, 200, 0, 100, 0)         ' base
    AssertAlmost 150, Geom_ValueToY(50, 200, -100, 100, 0)      ' zero no meio
    AssertAlmost 250, Geom_ValueToY(50, 200, 5, 5, 5)           ' escala degenerada
End Sub

Public Sub Test_Geom_Periods()
    AssertAlmost 3, Geom_Periods(2022, 2025, 1, 4)
    AssertAlmost 3, Geom_Periods("2022", "2025", 1, 4)
    AssertAlmost 2, Geom_Periods("T1", "T3", 1, 3)
    AssertAlmost 2, Geom_Periods(DateSerial(2020, 1, 1), DateSerial(2022, 1, 1), 1, 2), 0.01
    AssertAlmost 3, Geom_Periods(10, 40, 1, 4)                  ' numeros fora da faixa de anos: usa indices
End Sub

Public Sub Test_Geom_CategoryCenterX_NotBetween()
    AssertAlmost 100, Geom_CategoryCenterX(100, 300, 4, 1, False)    ' primeiro ponto na borda esquerda
    AssertAlmost 200, Geom_CategoryCenterX(100, 300, 4, 2, False)
    AssertAlmost 400, Geom_CategoryCenterX(100, 300, 4, 4, False)    ' ultimo na borda direita
    AssertAlmost 250, Geom_CategoryCenterX(100, 300, 1, 1, False)    ' n = 1: centro
End Sub

Public Sub Test_Geom_CategoryCenterY_And_ValueToX()
    AssertAlmost 125, Geom_CategoryCenterY(100, 200, 4, 1)     ' slot de 50: categoria 1 no topo, centro em 100+25
    AssertAlmost 275, Geom_CategoryCenterY(100, 200, 4, 4)
    AssertAlmost 50, Geom_ValueToX(50, 300, 0, 100, 0)
    AssertAlmost 350, Geom_ValueToX(50, 300, 0, 100, 100)
    AssertAlmost 200, Geom_ValueToX(50, 300, -100, 100, 0)
    AssertAlmost 50, Geom_ValueToX(50, 300, 5, 5, 5)           ' escala degenerada
End Sub
