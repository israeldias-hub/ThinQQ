VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestXlFake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Duble do Excel para os testes do suplemento (ligacao tardia: so os membros que modPptTqq usa). Run responde como um core com a API dada
' (Api Empty = core sem Core_PptApiVersion: erro 1004) e registra as macros chamadas. Nunca e um Excel de verdade.
Public Api As Variant
Public Calls As String                          ' "Core_Ping;Core_PptApiVersion;..." na ordem
Public LastEntries As String                    ' entradas recebidas por Core_PptSyncShapes
Public ShapesResult As String                   ' resposta de Core_PptSyncShapes ("" = "selecionados")
Public FailOn As String                         ' macro que levanta erro 1004
Public NoCore As Boolean                        ' Core_Ping nao responde "pong" (Excel sem o core)
Public AutoOn As Variant                        ' Core_PptGetFlag("autoon") (Empty = True)
Public DisplayAlerts As Boolean
Public InsertResult As String                   ' resposta de Core_PptInsertChart ("" = ok)
Public InsertAddsShape As Boolean               ' Core_PptInsertChart poe uma forma no slide recebido (como a colagem do core)
Public LastKind As String                       ' tipo recebido por Core_PptInsertChart
Public InsertSlideIndex As Long                 ' SlideIndex do slide recebido por Core_PptInsertChart
Public EditResult As String                     ' resposta de Core_PptEditData ("" = "Pasta - Excel")
Private mBooks As New Collection

Public Property Get Workbooks() As Collection
    Set Workbooks = mBooks
End Property

Public Function Run(macro As String, Optional a1 As Variant, Optional a2 As Variant, Optional a3 As Variant) As Variant
    Dim nm As String
    nm = Mid$(macro, InStr(macro, "!") + 1)
    Calls = Calls & nm & ";"
    If nm = FailOn Then Err.Raise 1004, "TestXlFake", "falha simulada em " & nm
    Select Case nm
        Case "Core_Ping": If NoCore Then Run = "" Else Run = "pong"
        Case "Core_PptGetFlag": If IsEmpty(AutoOn) Then Run = True Else Run = AutoOn
        Case "Core_PptApiVersion"
            If IsEmpty(Api) Then Err.Raise 1004, "TestXlFake", "macro inexistente"
            Run = Api
        Case "Core_PptSyncShapes"
            LastEntries = CStr(a2)
            If Len(ShapesResult) > 0 Then Run = ShapesResult Else Run = "selecionados"
        Case "Core_PptSyncPresentation": Run = "inteira"
        Case "Core_PptInsertChart"
            LastKind = CStr(a2)
            InsertSlideIndex = CLng(a1.SlideIndex)
            If InsertAddsShape And Len(InsertResult) = 0 Then a1.Shapes.AddShape 1, 10, 10, 50, 50
            Run = InsertResult
        Case "Core_PptEditData"
            LastEntries = CStr(a2)
            If Len(EditResult) > 0 Then Run = EditResult Else Run = "Pasta - Excel"
        Case "Core_PptDataSave": Run = 1
    End Select
End Function

Public Sub Quit()
End Sub
