VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPptEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Eventos do PowerPoint (suplemento ThinQQ.ppam). Nunca propagam erro. WindowActivate cobre a apresentacao que sai do Modo de Exibicao
' Protegido ("Habilitar Edicao") e a que ja estava aberta quando o suplemento carregou; PptTqq_OnOpen trata cada uma uma vez por sessao.
Public WithEvents App As Application

Private Sub App_AfterPresentationOpen(ByVal Pres As Presentation)
    On Error Resume Next
    PptTqq_OnOpen Pres
End Sub

Private Sub App_WindowActivate(ByVal Pres As Presentation, ByVal Wn As DocumentWindow)
    On Error Resume Next
    PptTqq_OnOpen Pres
End Sub
