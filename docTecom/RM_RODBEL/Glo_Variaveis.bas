Attribute VB_Name = "Glo_Variaveis"
Global sStatusMsg As String * 1 '0=ok,1=erro previsto,2-erro imprevisto, msg do sistema
Global sData As String * 10 ' data da ocorrencia
Global sHora As String * 5 'hora da ocorrencia
Global sTipo As String * 29 'tipo da movimentacao a ser gerada
Global sCodFun As String * 5 'codigo do funcionario
Global sMsg As String * 60 'Msg da ocorrencia

Type ArqTexto
     Texto As String * 110
     FFinal As String * 2
End Type
Global sTexto As ArqTexto

Rem aquivo de Impressao em disco PARA EXEL
Type ArqImpressaoE
     FCampo1000 As String * 1000
     FFinal As String * 2
End Type
Global Arq_ImpressaoE As ArqImpressaoE

Rem aquivo de Impressao em disco
Type ArqImpressao
     FCampo136 As String * 136
     FFinal As String * 2
End Type
Global Arq_Impressao As ArqImpressao

