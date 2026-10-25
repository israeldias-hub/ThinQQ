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
' Protegido ("Habilitar Edicao"); PptTqq_OnOpen trata cada uma uma vez enquanto aberta (so com janela num PowerPoint visivel).
Public WithEvents App As Application

Private Sub App_AfterPresentationOpen(ByVal Pres As Presentation)
    On Error Resume Next
    PptTqq_OnOpen Pres
End Sub

Private Sub App_WindowActivate(ByVal Pres As Presentation, ByVal Wn As DocumentWindow)
    On Error Resume Next
    PptTqq_OnOpen Pres
    PptLinks_OnSelChange                    ' 1.11.0: o dialogo Vinculos aberto passa a mostrar esta apresentacao
End Sub

' Selecao mudou: o estado de "Atualizar automaticamente" segue o grafico selecionado. So invalida a faixa e, com o dialogo Vinculos aberto,
' o atualiza (so etiquetas; nada de MsgBox nem Excel).
Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    On Error Resume Next
    PptRx_InvalidateElem
    PptLinks_OnSelChange False              ' so marca as linhas (sem reler o deck a cada clique; revisao final M4)
End Sub

' Outro slide (painel de miniaturas, classificacao de slides): o dialogo Vinculos aberto acompanha (1.11.0). So etiquetas.
Private Sub App_SlideSelectionChanged(ByVal SldRange As SlideRange)
    On Error Resume Next
    PptLinks_OnSelChange
End Sub

' Salvar a apresentacao tambem salva a pasta de dados dos graficos criados pelo PowerPoint (1.8.0). Sonda PG5: este evento chega DEPOIS
' da gravacao do .pptx (e, no Salvar como, com o nome novo); o xl.Run que salva a pasta de dados funciona aqui.
Private Sub App_PresentationSave(ByVal Pres As Presentation)
    On Error Resume Next
    PptTqq_OnSave Pres
End Sub

' Fechada de fato (depois de "Salvar?"): esquece a apresentacao, para a proxima abertura perguntar de novo.
Private Sub App_PresentationCloseFinal(ByVal Pres As Presentation)
    On Error Resume Next
    PptTqq_OnClose Pres
End Sub
