VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestPptTextFake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Duble de uma caixa de texto do PowerPoint para o texto vinculado (modPptText), sem PowerPoint: a mesma instancia faz o papel de forma,
' TextFrame e TextRange; Characters(inicio, tamanho) devolve um trecho (outra instancia ligada a raiz) cujo Text troca so aqueles caracteres.
' TextWrites conta as gravacoes do texto (prova que a sincronizacao sem mudanca nao reescreve); FailTextWrite faz a proxima gravacao falhar.
Public Tags As Object                           ' TestPptTags
Public HasTextFrame As Long                     ' -1 = msoTrue
Public TextWrites As Long
Public FailTextWrite As Boolean
Private mText As String
Private mRoot As TestPptTextFake
Private mSt As Long
Private mLn As Long

Public Sub InitRun(root As TestPptTextFake, st As Long, ln As Long)
    Set mRoot = root: mSt = st: mLn = ln
End Sub

Public Property Get TextFrame() As Object
    Set TextFrame = Me
End Property

Public Property Get TextRange() As Object
    Set TextRange = Me
End Property

Public Property Get Text() As String
    If mRoot Is Nothing Then Text = mText Else Text = Mid$(mRoot.Text, mSt, mLn)
End Property

Public Property Let Text(ByVal v As String)
    If mRoot Is Nothing Then
        If FailTextWrite Then FailTextWrite = False: Err.Raise vbObjectError + 9993, "TestPptTextFake", "falha injetada (texto)"
        mText = v
        TextWrites = TextWrites + 1
    Else
        mRoot.Text = Left$(mRoot.Text, mSt - 1) & v & Mid$(mRoot.Text, mSt + mLn)
        mLn = Len(v)
    End If
End Property

' Texto posto pelo teste (nao conta como gravacao).
Public Sub Preset(ByVal v As String)
    mText = v
End Sub

Public Function Characters(ByVal st As Long, ByVal ln As Long) As Object
    Dim r As TestPptTextFake
    Set r = New TestPptTextFake
    If mRoot Is Nothing Then r.InitRun Me, st, ln Else r.InitRun mRoot, mSt + st - 1, ln
    Set Characters = r
End Function

Public Sub InsertBefore(ByVal v As String)
    If mRoot Is Nothing Then Text = v & mText Else mRoot.Text = Left$(mRoot.Text, mSt - 1) & v & Mid$(mRoot.Text, mSt)
End Sub

Public Sub InsertAfter(ByVal v As String)
    If mRoot Is Nothing Then Text = mText & v Else mRoot.Text = Left$(mRoot.Text, mSt + mLn - 1) & v & Mid$(mRoot.Text, mSt + mLn)
End Sub
