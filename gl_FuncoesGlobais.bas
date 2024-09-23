Attribute VB_Name = "gl_FuncoesGlobais"
Option Explicit

Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public sEMP_NOME As String * 25 ' NOME DA EMPRESA
Public sNome_Imp As String ' nome da impressora padrao para impressao
Public nSel_Imp As Integer ' caso esteja c/0 nao aparece a selecao caso sim, aparece tela de selecao de impressora
Public nTipo_Imp As Integer ' caso esteja c/0 a impressora é grafica, c/1 sera entendida como genérica
Public nTamCodFunc As Integer 'tamanho do capo do codigo de funcionario
Public nTamCodPosto As Integer 'tamanho do capo do codigo do posto
'Public Const DRIVE_CDROM = 5
'Public Const DRIVE_FIXED = 3
'Public Const DRIVE_RAMDISK = 6
'Public Const DRIVE_REMOTE = 4
'Public Const DRIVE_REMOVABLE = 2
Public sNome_Usuario As String 'nome do usuario

Global sNomeEmpresa As String ' nome da empresa que se esta logado
Global sNomeBanco As String ' nome do banco que se esta logado
Global sBancoRM As String ' nome do banco que se esta logado para o RM
Global sBancoRodbel As String 'nome do banco que se esta logado para o Rodbel
Global sBancoUnimed As String 'nome do banco que se esta logado para a movimentação de impor. e expor. Unimed

Global sUsuario As String ' codigo do usuario em operacao
Global sNomeUsuario As String * 15 'nome do úsuario
Global sCodEempresa As String ' codigo da empresa em operacao
Global nLinhasImp As Integer ' numero de linhas impressas
Global nTotalLinhaImp As Integer ' numero de linhas impressas no formulario
Global nDiasManArmas As Integer ' numero de dias para manutenção das armas
Global ADOConnection As ADODB.Connection
Global nAtualizadas As Double 'contador geral
Public rRec_cliente As ADODB.Recordset 'contem os clientes da tabela de etiquetas
'--------------------------------------------------------------------
Type ArquivoTexto ' Esta variavel servirá para leitura do arquivo que contera as informacoes do banco e as empresas
     LinhaTexto As String * 80
End Type

Global ARQUIVO_TEXTO As ArquivoTexto

'--------------------------------------------------------------------
Rem arquivo para movimentacao de ocorrencias para a folha de pagto
Type ArqMovInvetario
     FRegistro As String * 10
     FTipo As String * 1
     FFinal As String * 2
End Type

Global Arq_Mov_Inventario As ArqMovInvetario
'--------------------------------------------------------------------
Rem aquivo de Impressao em disco
Type ArqImpressao
     FCampo136 As String * 136
     FFinal As String * 2
End Type
Global Arq_Impressao As ArqImpressao
'----------------------------------------------------------
Rem ARQUIVO DE EXPORTACAO PARA O R3 ATUALIZAR OS PALETS NAS ETIQUETAS
Type Arq_R3
     Fetiqueta As String * 10
     Fpallet As String * 20
     FFinal As String * 2
End Type
Global ArqR3 As Arq_R3
'----------------------------------------------------------
Rem ARQUIVO DE EXPORTACAO/IMPORTACAO DO MOVIMENTO DA UNIMED PARA O RM-LABORE
Type Arq_Mov_Unimed
     Fchapa As String * 16
     Fdtpagto As String * 8
     Fcodevento As String * 4
     Fhora As String * 6
     Frefer As String * 15
     Fvalor As String * 15
     Fvaloreal As String * 15
     Falterado As String * 1
     FFinal As String * 2
End Type
Global ArqMovUnimed As Arq_Mov_Unimed

Public Function RecordAdodb(cRecAdo As ADODB.Recordset) As ADODB.Recordset

Rem record set do adodb dinamico
Dim cRec As ADODB.Recordset
Dim nx As Integer
Dim Nz As Integer
Dim sType As Integer

On Error GoTo Erro

Set cRec = New ADODB.Recordset
   
   
For nx = 0 To cRecAdo.Fields.Count - 1
    sType = cRecAdo.Fields(nx).Type
    If sType > 21 Then
       sType = GetDataTypeEnum(cRecAdo.Fields(nx).Type)
    End If
    cRec.Fields.Append cRecAdo.Fields(nx).Name, sType
Next

cRec.Open
If cRecAdo.RecordCount = 0 Then
   Set RecordAdodb = cRec
   Exit Function
End If

cRecAdo.MoveFirst
   
While Not cRecAdo.EOF
      
      cRec.AddNew
      For nx = 0 To cRec.Fields.Count - 1
          If cRecAdo(nx).Type > 0 And cRecAdo(nx).Type < 6 Then
             cRec.Fields.Item(nx).Value = IIf(IsNull(cRecAdo(nx).Value), 0, Trim(cRecAdo(nx).Value))
          Else
             cRec.Fields.Item(nx).Value = IIf(IsNull(cRecAdo(nx).Value), " ", Trim(cRecAdo(nx).Value))
          End If
      Next
      cRec.Update
      cRecAdo.MoveNext
Wend

Set RecordAdodb = cRec

Exit Function

Erro:

MsgBox ""

End Function
Function GetDataTypeEnum(lngDataTypeEnum As Long) As Integer
      'Given ADO data-type constant, returns readable constant name.
      Dim strReturn As String
         Select Case lngDataTypeEnum
            Case 0: strReturn = 0 ' "adEmpty"
            Case 2: strReturn = 2 ' "adSmallInt"
            Case 3: strReturn = 3 ' "adInteger"
            Case 4: strReturn = 4 ' "adSingle"
            Case 5: strReturn = 5 ' "adDouble"
            Case 6: strReturn = 6 ' "adCurrency"
            Case 7: strReturn = 7 ' "adDate"
            Case 8: strReturn = 8 ' "adBSTR"
            Case 9: strReturn = 9 ' "adIDispatch"
            Case 10: strReturn = 10 ' "adError"
            Case 11: strReturn = 11 ' "adBoolean"
            Case 12: strReturn = 12 ' "adVariant"
            Case 13: strReturn = 13 ' "adIUnknown"
            Case 14: strReturn = 14 ' "adDecimal"
            Case 16: strReturn = 16 ' "adTinyInt"
            Case 17: strReturn = 17 ' "adUnsignedTinyInt"
            Case 18: strReturn = 18 ' "adUnsignedSmallInt"
            Case 19: strReturn = 19 ' "adUnsignedInt"
            Case 20: strReturn = 20 ' "adBigInt"
            Case 21: strReturn = 21 ' "adUnsignedBigInt"
            
            Case 131: strReturn = "adNumeric"
            Case 132: strReturn = "adUserDefined"
            Case 72: strReturn = "adGUID"
            Case 133: strReturn = "adDBDate"
            Case 134: strReturn = "adDBTime"
            Case 135: strReturn = 8 ' "adDBTimeStamp"
            Case 129: strReturn = "adChar"
            Case 200: strReturn = "adVarChar"
            Case 201: strReturn = "adLongVarChar"
            Case 130: strReturn = "adWChar"
            Case 202: strReturn = 8 ' "adVarWChar"
            Case 203: strReturn = "adLongVarWChar"
            Case 128: strReturn = "adBinary"
            Case 204: strReturn = "adVarBinary"
            Case 205: strReturn = "adLongVarBinary"
         Case Else:
            strReturn = "Unknown DataTypeEnum of " & lngDataTypeEnum _
             & " found."
         End Select
         GetDataTypeEnum = strReturn
      End Function
Public Function Retorna_DiaSemana(ByVal dDATA As String) As String

If VBA.Weekday(CDate(dDATA)) = 1 Then Retorna_DiaSemana = "DOMINGO": Exit Function
If VBA.Weekday(CDate(dDATA)) = 2 Then Retorna_DiaSemana = "SEGUNDA-FEIRA": Exit Function
If VBA.Weekday(CDate(dDATA)) = 3 Then Retorna_DiaSemana = "TERCA-FEIRA": Exit Function
If VBA.Weekday(CDate(dDATA)) = 4 Then Retorna_DiaSemana = "QUARTA-FEIRA": Exit Function
If VBA.Weekday(CDate(dDATA)) = 5 Then Retorna_DiaSemana = "QUINTA-FEIRA": Exit Function
If VBA.Weekday(CDate(dDATA)) = 6 Then Retorna_DiaSemana = "SEXTA-FEIRA": Exit Function
If VBA.Weekday(CDate(dDATA)) = 7 Then Retorna_DiaSemana = "SABADO": Exit Function

End Function
Public Function Retorna_Mes(ByVal dMes As Integer) As String

If dMes = 1 Then Retorna_Mes = "JANEIRO": Exit Function
If dMes = 2 Then Retorna_Mes = "FEVEREIRO": Exit Function
If dMes = 3 Then Retorna_Mes = "MARCO": Exit Function
If dMes = 4 Then Retorna_Mes = "ABRIL": Exit Function
If dMes = 5 Then Retorna_Mes = "MAIO": Exit Function
If dMes = 6 Then Retorna_Mes = "JUNHO": Exit Function
If dMes = 7 Then Retorna_Mes = "JULHO": Exit Function
If dMes = 8 Then Retorna_Mes = "AGOSTO": Exit Function
If dMes = 9 Then Retorna_Mes = "SETEMBRO": Exit Function
If dMes = 10 Then Retorna_Mes = "OUTUBRO": Exit Function
If dMes = 11 Then Retorna_Mes = "NOVEMBRO": Exit Function
If dMes = 12 Then Retorna_Mes = "DEZEMBRO": Exit Function

End Function
Function Proximo_Mes(ByVal dDate As Date) As Date
Dim nMes As Integer
Dim sdata As String
Dim nano As Integer
Dim sDias As String

sDias = "312931303130313130313031"

sdata = dDate
nMes = Month(dDate)
nMes = nMes + 1
nano = Val(Mid$(sdata, 7, 4))

If nMes = 13 Then
   nMes = 1
   nano = nano + 1
End If

If Val(Mid$(sdata, 1, 2)) >= 28 And Val(Mid$(sdata, 1, 2)) <= 31 Then
   sdata = Mid$(sDias, (nMes * 2) - 1, 2) & "/" & Format(nMes, "00") & "/" & Format(nano, "0000")
Else
   sdata = Mid$(sdata, 1, 3) & Format(nMes, "00") & "/" & Format(nano, "0000")
End If

If Mid$(sdata, 4, 2) = "02" Then
   If IsDate(sdata) Then
      Proximo_Mes = sdata
   Else
      Proximo_Mes = Replace(sdata, "29", "28")
   End If
Else
   Proximo_Mes = sdata
End If



End Function
Public Function Testa_Numerico(ByVal sCampo As String, ByVal nTamanhoCampo As Integer) As Boolean
Dim nx As Integer

Testa_Numerico = True

For nx = 1 To nTamanhoCampo
    If Asc(Mid$(sCampo, nx, 1)) < 48 Or _
       Asc(Mid$(sCampo, nx, 1)) > 57 Then
       Testa_Numerico = False
       Exit For
    End If
Next

End Function

Public Function Cripta(cTexto As String) As String
     Dim cTemp As String
     Dim x As Integer
     cTemp = ""
     For x = 1 To Len(cTexto)
        cTemp = cTemp + Mid(cTexto, x, 1)
     Next
     Cripta = ""
     For x = 1 To Len(cTemp)
        Cripta = Cripta + Chr(256 - Asc(Mid(cTemp, x, 1)))
     Next
End Function

Public Function UnCripta(cTexto As String) As String
     Dim cTemp As String
     Dim x As Integer
     cTemp = ""
     For x = 1 To Len(cTexto)
        cTemp = cTemp + Mid(cTexto, x, 1)
     Next
     UnCripta = ""
     For x = 1 To Len(cTemp)
        UnCripta = UnCripta + Chr(256 - Asc(Mid(cTemp, x, 1)))
     Next
End Function
Public Function CCTempneLog() As neLog
     Set CCTempneLog = New neLog
End Function
Public Function CCTempneFormulario() As neFormulario
     Set CCTempneFormulario = New neFormulario
End Function
Public Function CCTempneUsuario() As neUsuario
      Set CCTempneUsuario = New neUsuario
End Function
Public Function CCTempneUniMvFun() As neUniMvFun
      Set CCTempneUniMvFun = New neUniMvFun
End Function
Public Function CCTempConect() As daAbertura
     Set CCTempConect = New daAbertura
End Function
Public Function CCTempneUniColigada() As neUniColigada
     Set CCTempneUniColigada = New neUniColigada
End Function
Public Function CCTempneTabRegPagto() As neTabRegPagto
     Set CCTempneTabRegPagto = New neTabRegPagto
End Function

'''Um objeto NODE é um item do treeview. Este item pode ou não conter sub-itens. Todos os objetos NODE do TreeView são guardados dentro de uma coleção NODE, que é exposta na forma de uma propriedade do componente.
'''
'''Assim sendo, através da coleção NODE realizamos o acesso a qualquer item do treeview.
'''
'''É através da coleção NODE que iremos realizar a inclusão de um item no TreeView. O método ADD da coleção NODE possuí os seguintes parâmetros:
'''
'''Relative: A adição de um nó pode ser feita em uma posição relativa a partir de um nó já existente.
'''
'''Relashionship: Este parâmetro identifica qual a relação entre o nó que está sendo adicionado e o nó já existente. Existem 05 opções:
'''
'''01. tvwFirst: O novo nó deverá ser incluido como o 1o nó no nível em que está o nó indicado em Relative.
'''
'''02. tvwLast: O novo nó deverá ser incluido como o último nó no nível em que está o nó indicado em Relative.
'''
'''03. tvwPrevious: O novo nó deverá ser incluido na posição imediatamente anterior a posição do nó indicado em Relative.
'''
'''04. tvwNext: O novo nó deverá ser incluido na posição imediatamente posterior a posição do nó indicado em Relative.
'''
'''05. tvwChild: O novo nó deverá ser incluido como Child. Ou seja, como um sub-nó, do nó indicado em Relative.
'''
'''Key: É uma palavra chave, uma espécie de código, que identifica o nó de forma única. A key é opcional, mas pode ajudar bastante na hora de localizar uma determinada informação dentro da TreeView.
'''
'''Text: O texto em si do nó.
'''
'''Image: A imagem que deverá ser utilizada para representar o nó.
'''
'''Selected Image: A imagem que deverá ser utilizada para representar o nó quando este estiver selecionado.
'''
'''Vejamos, enfim, um pequeno exemplo da adição de nós:
'''
'''TreeView1.Nodes.Add , , "iMasters", "iMasters"
'''TreeView1.Nodes.Add , , "Terra", "Terra"
'''TreeView1.Nodes.Add , , "Google", "Google"
'''
'''Com estas 03 instruções teremos adicionado nós referentes. Vamos agora adicionar, para cada um dos Nodes, 03 novos Nodes.
'''
'''TreeView1.Nodes.Add "iMasters", tvwChild, "Programação", "Compras"
'''TreeView1.Nodes.Add "iMasters", tvwChild, "Cursos", "Vendas"
'''TreeView1.Nodes.Add "iMasters", tvwChild, "Pesquisa", "Diretoria"
'''TreeView1.Nodes.Add "Terra", tvwChild, "Chat", "Compras"
'''TreeView1.Nodes.Add "Terra", tvwChild, "Email", "Vendas"
'''TreeView1.Nodes.Add "Terra", tvwChild, "Notícias", "Diretoria"
'''TreeView1.Nodes.Add "Google", tvwChild, "Vendas", "Compras"
'''TreeView1.Nodes.Add "Google", tvwChild, "Empregos", "Vendas"
'''TreeView1.Nodes.Add "Google", tvwChild, "Vagas", "Diretoria"
'''
'''Observe que no item Relative foi possível especificar a Key fornecida para os primeiros itens. Se não tivessemos incluido uma KEY teríamos que nos referir ao nó por sua posição numérica.
'''
'''Essas instruções de criação de nós devem ser inseridas no form_load da aplicação. Pronto, nossa aplicação já esta funcionando e agora você já sabe utilizar o TreeView.

