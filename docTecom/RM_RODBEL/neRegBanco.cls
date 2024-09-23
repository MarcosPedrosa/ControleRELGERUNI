VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neRegBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DO ATUALIZACAO DOS FUNCIONARIOS NOVATOS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionarios_Novatos(ByVal sBdRm As String, _
                                     ByVal sBdRodBel As String, _
                                     ByVal c_Data_Admissao As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset 'record set que pesquisara se existe o funcionario no Rodbel
Dim nx As Integer 'VARIAVEL CONTAGEM LOOP
Dim nAre_Numero As Integer 'CAMPO PARA GRAVACAO DO ATRIBUTO DA TABELA CRED-ACESSO

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura

Rem  instancia a conecao com o banco
Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

Set cFields = New Collection
cFields.Add c_Data_Admissao, "c_Data_Admissao"

Set Funcionarios_Novatos = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 1)

nGravou = 0
nLidos = 0

If Funcionarios_Novatos.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR MOVIMENTACAO NO BANCO RODBEL CASO NAO EXISTA NO CADASTRO
   
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionarios_Novatos.RecordCount

   Funcionarios_Novatos.MoveFirst
   
   While Not Funcionarios_Novatos.EOF
'         If Val(Funcionarios_Novatos!CHAPA) = 3284 Then
'             nGrvOk = 0
'         End If
         Rem zerar a variavel para 0, se oiuver atualizacao sera trocado por 1, de acordo com as criticas abaixo.
         nGrvOk = 0
         
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE RELACIONAMENTO CRED_PESSOAS e
         Rem PESSOAS, CASO EXISTA, IGNORAR REGISTRO
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Novatos!CHAPA
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)
         
         Rem CASO EXISTA REGISTRO, PROXIMO CANDIDATO
         If rRs.RecordCount > 0 Then
            GoTo PROXIMO_REGISTRO
         End If
         
         Rem VERIFICA A EXISTENCIA DO CENTRO DE CUSTO DA RM NO BANCO DA RODBEL,
         Rem RELACIONANDO O CAMPO INTEGRGERENCIAL DA RM COM A TABELA "ESTRUTURA", NA RODBEL
         
         Set cFields = New Collection
         cFields.Add IIf(IsNull(Funcionarios_Novatos!INTEGRGERENCIAL), "0", Replace(Funcionarios_Novatos!INTEGRGERENCIAL, ".", ""))
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Estrutura(ADOConnectRodbel, cFields, 6)
         
         If rRs!ENCONTRADO = 0 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Novatos"
            sCodFun = Funcionarios_Novatos!CHAPA
            sMsg = "Funcionario com o Centro de custo inexistente -  " & cFields(1)
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            nGrvOk = 1
         End If
         
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. PESSOA, CASO EXISTA, ERRO.
         Set cFields = New Collection
         cFields.Add Funcionarios_Novatos!CHAPA
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 2)
         
         If rRs.RecordCount = 1 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Novatos"
            sCodFun = Funcionarios_Novatos!CHAPA
            sMsg = "Existe TAB.CREDENCIAIS e Nao Existe na TAB.PESSOAS"
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            nGrvOk = 1
         
         End If
         
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE RELACIONAMENTO CRED_PESSOAS, CASO EXISTA, ERRO.
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Novatos!CHAPA
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 7)
         
         If rRs.RecordCount > 0 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Novatos"
            sCodFun = Funcionarios_Novatos!CHAPA
            sMsg = "Existe na TABS. CRED_PESSOA E NÃO EXISTE NO PESSOAS"
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            nGrvOk = 1
         
         End If
         
         If nGrvOk = 0 Then
            
            Rem ************************************************************************************************
            Rem ********* GRAVAR AS TABELAS DE PESSOAS, CREDENCIAIS E CRED_PESSOAS *****************************
            Rem ************************************************************************************************
            
            Rem *************  PESSOAS  ***************
            Set cFields = New Collection
            cFields.Add Funcionarios_Novatos!CHAPA
            cFields.Add IIf(IsNull(Funcionarios_Novatos!NOME), "", Funcionarios_Novatos!NOME)
            cFields.Add IIf(IsNull(Funcionarios_Novatos!CPF), "", Funcionarios_Novatos!CPF)
            cFields.Add "" 'OBSERVACAO
            cFields.Add "" 'E-MAIL
            cFields.Add IIf(IsNull(Funcionarios_Novatos!CARTIDENTIDADE), "", Funcionarios_Novatos!CARTIDENTIDADE)
            cFields.Add IIf(IsNull(Funcionarios_Novatos!INTEGRGERENCIAL), "", Replace(Funcionarios_Novatos!INTEGRGERENCIAL, ".", ""))
            cFields.Add "0" 'FLAGREENTRADA
            cFields.Add "0" 'MASTER
            cFields.Add "Nao_Disponivel.JPG" 'FOTO
            cFields.Add "" 'SENHA
            cFields.Add " " 'RAMAL
            cFields.Add " " 'CAMPOPERSO1
            cFields.Add " " 'CAMPOPERSO2
            cFields.Add " " 'CAMPOPERSO3
            cFields.Add " " 'CAMPOPERSO4
            cFields.Add " " 'CAMPOPERSO5
            cFields.Add "8" 'PESSIT_NUMERO - VER QUE MUDANCA VAI SER
            cFields.Add "" 'DTSITINICIAL
            cFields.Add "" 'DTSITFINAL
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'TOK_NUMEROSERIE
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 3)
            
            Rem ************************************************************************************************
            Rem NAO GERADO PARA A TECOM, POIS OS NUMEROS DOS CRACHAS ESTAO CADASTRADOS NO SISTEMA E SERÃO DEPOIS ASSOCIADOS
            Rem AOS FUNCIONARIOS CADASTRADOS

'''            Rem *************  CREDENCIAIS  ***************
'''            Set cFields = New Collection
'''            cFields.Add Funcionarios_Novatos!CHAPA 'CRED_NUMERO
'''            cFields.Add "0" 'CRED_FLAGREENTRADA
'''            cFields.Add "0" 'CRED_AREANOMOMENTO
'''            cFields.Add "" 'CRED_DTVALIDINICIAL
'''            cFields.Add "" 'CRED_DTVALIDFINAL
'''            cFields.Add "0" 'CRED_BLOQUEADA
'''            cFields.Add "0" 'CRED_FLAGCATRACA
'''            cFields.Add "1" 'CRTI_NUMERO
'''            cFields.Add "0" 'CRED_MASTER
'''            cFields.Add "" 'CRED_ULTPASSAGEM
'''            cFields.Add "" 'CRED_EQPINUMERO
'''            cFields.Add " " 'CRED_SENHA
'''            cFields.Add "0" 'CRED_PROV
'''            cFields.Add "1" 'CRED_PROVDEVOLV
'''            cFields.Add "0" 'CRED_NUMRA
'''
'''            Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 4)
'''
'''            Rem ************************************************************************************************
'''
'''            Rem *************  CRED_PESSOAS  ***************
'''            Set cFields = New Collection
'''            cFields.Add Funcionarios_Novatos!CHAPA 'CRED_NUMERO
'''            cFields.Add Funcionarios_Novatos!CHAPA  'PES_NUMERO
'''            cFields.Add Format(Now(), "YYYY") & Format(Now(), "MM") & Format(Now(), "DD") 'CRPES_DATAENTRADA
'''            cFields.Add "99991231" 'CRPES_DATASAIDA
'''
'''            Call RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 5)
         
            Rem ************************************************************************************************
'''            nGravou = nGravou + 1
'''
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Novatos"
            sCodFun = Funcionarios_Novatos!CHAPA
            Rem modificar a data abaixo, esta errada
            sMsg = "Inclusão Novatos, Dt.Admissao - " & c_Data_Admissao
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg

            Call Gerar_Situacao_Log(cFields)
'''
'''            Rem Aqui começa a outra parte da atualizacao dos novatos, Criar os acessos nos relogios eletronicos
'''
'''            Rem LEITURA DA TABELA DE GRUPOS_AREAS, VERIFICANDO OS RELOGIOS E SUAS EXISTENCIAS
'''
'''            Set cFields = New Collection
'''            cFields.Add Funcionarios_Novatos!CHAPA
'''            Set rRs = New ADODB.Recordset
'''            Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 18)
'''
'''            If rRs.RecordCount = 0 Then
'''               ADOConnectRodbel.CommitTrans
'''               ADOConnectRodbel.BeginTrans
'''               GoTo PROXIMO_REGISTRO
'''            End If
'''
'''            Rem VERIFICAR OS REGISTROS DE GRUPOS-AREAS SELECIONADOS E GRAVAR OS REGISTROS NA TABELA DE CRED_ACESSO
'''
'''            While Not rRs.EOF
'''                  nGrvOk = 0
'''
'''                  If rRs!GRP_NUMERO = 1 Then
'''                     Rem aqui refere-se as funcoes que nao serao incluidas, pois se referem a gerentes e diretores com passes livres.
'''                     If Funcionarios_Novatos!CARGO <> 24 And _
'''                        Funcionarios_Novatos!CARGO <> 27 And _
'''                        Funcionarios_Novatos!CARGO <> 32 And _
'''                        Funcionarios_Novatos!CARGO <> 72 And _
'''                        Funcionarios_Novatos!CARGO <> 92 And _
'''                        Funcionarios_Novatos!CARGO <> 200 And _
'''                        Funcionarios_Novatos!CARGO <> 203 And _
'''                        Funcionarios_Novatos!CARGO <> 489 And _
'''                        Funcionarios_Novatos!CARGO <> 551 Then
'''                        nGrvOk = 1
'''                     End If
'''                  End If
'''
'''                  Rem GRAVAR REGISTRO DE CRED-ACESSO DE ACORDO COM A REGRA ACIMA
'''
'''                  If nGrvOk = 0 Then
'''                     nAre_Numero = rRs!ARE_NUMEROENTRADA
'''
'''                     For Nx = 1 To 2
'''                         Rem *************  CRED-ACESSO  ***************
'''                         Set cFields = New Collection
'''                         cFields.Add "'" & Funcionarios_Novatos!CHAPA & "'" 'CRED_NUMERO
'''                         cFields.Add rRs!GRP_NUMERO 'GRP_NUMERO
'''                         cFields.Add nAre_Numero 'ARE_NUMERO
'''                         cFields.Add "1" 'CRACTI_NUMERO
'''                         cFields.Add "" 'CRAC_DIFPORGRUPO
'''                         cFields.Add "" 'JORN_NUMERO
'''                         cFields.Add "" 'TURN_NUMERO
'''                         cFields.Add "999" 'CRAC_QTDTOTALACESSOS
'''                         cFields.Add "" 'CRAC_QTDACESSOS
'''                         cFields.Add "0" 'CRAC_VERFSENHA
'''                         cFields.Add "0" 'CRAC_VERFDIGITAL
'''                         cFields.Add "01/01/2010" 'CRAC_ULTPASSAGEM
'''
'''                         Call RegBanco_Critica_Cred_Acesso(ADOConnectRodbel, cFields, 19)
'''                         nAre_Numero = rRs!ARE_NUMEROSAIDA
'''
'''                     Next
'''
'''                  End If
'''
'''                  rRs.MoveNext
'''
'''            Wend
'''
'''
'''
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionarios_Novatos.EOF Then
            Funcionarios_Novatos.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If




If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Novatos"
   sCodFun = "0000"
   sMsg = "Nao ha Novatos no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If


ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
   sCodFun = Funcionarios_Novatos!CHAPA
Else
   sStatusMsg = "2"
   sCodFun = "0000"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Novatos"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DO ATUALIZACAO DOS FUNCIONARIOS DESLIGADOS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionarios_Desligados(ByVal sBdRm As String, _
                                        ByVal sBdRodBel As String, _
                                        ByVal c_Data_Demissao As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

cFields.Add c_Data_Demissao, "c_Data_Demissao"

Set Funcionarios_Desligados = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 9)

nGravou = 0
nLidos = 0

If Funcionarios_Desligados.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR MOVIMENTACAO NO BANCO RODBEL CASO NAO EXISTA NO CADASTRO
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionarios_Desligados.RecordCount

   Funcionarios_Desligados.MoveFirst
   
   While Not Funcionarios_Desligados.EOF
'         If Val(Funcionarios_Desligados!CHAPA) = 1429 Then
'             nGrvOk = 0
'         End If
         
         nGrvOk = 0
         
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Desligados!CHAPA
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE RELACIONAMENTO CRED_PESSOAS e PESSOAS,
         Rem CASO EXISTA, VERIFICAR SE JA FOI DEMITIDO CASO SIM IGNORAR REGISTRO
         
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)
         If rRs.RecordCount = 0 Then
            GoTo PROXIMO_REGISTRO
         Else
            If rRs!pessit_numero = 11 Then
               GoTo PROXIMO_REGISTRO
            End If
         End If
         
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. PESSOAS, NAO CASO NAO EXISTA, ERRO.
         Rem SELECIONAR O REGISTRO PESSOAS DA RODBEL PARA VERIFICAR O STATUS DO CAMPO PESSIT_NUMERO E
         Rem CRED_BLOQUEADA DA TABELA CREDENCIAIS, E TESTAR A SITUACAO QUE VEIO DA RM
         Set cFields = New Collection
         cFields.Add Funcionarios_Desligados!CHAPA
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)
         
         If rRs.RecordCount = 0 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Desligados"
            sCodFun = Funcionarios_Desligados!CHAPA
            sMsg = "Funcionarios nao Existe na TAB.CREDENCIAIS"
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            nGrvOk = 1
         
         End If
         
         If nGrvOk = 0 Then
            
            Rem ************************************************************************************************
            Rem ********* GRAVAR AS TABELAS DE PESSOAS, CREDENCIAIS ********************************************
            Rem ************************************************************************************************
            
            Rem *************  PESSOAS  ***************
            Rem  ATUALIZACAO DO CAMPO PESSIST_NUMERO sStatusMsg = 3
            Set cFields = New Collection
            cFields.Add Funcionarios_Desligados!CHAPA
            cFields.Add "" 'NOME
            cFields.Add "" 'CPF
            cFields.Add "" 'OBSERVACAO
            cFields.Add "" 'E-MAIL
            cFields.Add "" 'IDENTIDADE
            cFields.Add "" 'CENTRO DE CUSTO
            cFields.Add "" 'FLAGREENTRADA
            cFields.Add "" 'MASTER
            cFields.Add "" 'FOTO
            cFields.Add "" 'SENHA
            cFields.Add "" 'RAMAL
            cFields.Add "" 'CAMPOPERSO1
            cFields.Add "" 'CAMPOPERSO2
            cFields.Add "" 'CAMPOPERSO3
            cFields.Add "" 'CAMPOPERSO4
            cFields.Add "" 'CAMPOPERSO5
            cFields.Add "11" 'PESSIT_NUMERO
            cFields.Add "" 'DTSITINICIAL
            cFields.Add "" 'DTSITFINAL
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'TOK_NUMEROSERIE
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 10)
            
            Rem ************************************************************************************************
            
            Rem *************  CREDENCIAIS  ***************
            rRs.MoveFirst
            While Not rRs.EOF
                Set cFields = New Collection
                cFields.Add rRs!CRED_NUMERO 'CRED_NUMERO
                cFields.Add "" 'CRED_FLAGREENTRADA
                cFields.Add "" 'CRED_AREANOMOMENTO
                cFields.Add "" 'CRED_DTVALIDINICIAL
                cFields.Add "" 'CRED_DTVALIDFINAL
                cFields.Add "1" 'CRED_BLOQUEADA
                cFields.Add "" 'CRED_FLAGCATRACA
                cFields.Add "" 'CRTI_NUMERO
                cFields.Add "" 'CRED_MASTER
                cFields.Add "" 'CRED_ULTPASSAGEM
                cFields.Add "" 'CRED_EQPINUMERO
                cFields.Add "" 'CRED_SENHA
                cFields.Add "" 'CRED_PROV
                cFields.Add "" 'CRED_PROVDEVOLV
                cFields.Add "" 'CRED_NUMRA
                
                Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 11)
                rRs.MoveNext
            Wend
            Rem ************************************************************************************************
            
            nGravou = nGravou + 1
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Desligados"
            sCodFun = Funcionarios_Desligados!CHAPA
            sMsg = "Desligamento funcionario, no Periodo de " & c_Data_Demissao
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionarios_Desligados.EOF Then
            Funcionarios_Desligados.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Desligados"
   sCodFun = "0000"
   sMsg = "Nao ha Desligamento no periodo de " & c_Data_Demissao
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

'If nGravou > 0 Then
'   If Err.Number > 9999 And Err.Number < 10001 Then 'Erro de consistencia de dados para atualizacao do RodBel
'
'   Else
'
'   End If
'End If

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Desligados"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

'ADOConnectRm.RollbackTrans
'ADOConnectRodbel.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DO ATUALIZACAO DOS FUNCIONARIOS EM FERIAS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionarios_Ferias(ByVal sBdRm As String, _
                                    ByVal sBdRodBel As String, _
                                    ByVal c_Data_Ferias As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

cFields.Add c_Data_Ferias, "c_Data_Ferias"

Set Funcionarios_Ferias = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 12)

nGravou = 0
nLidos = 0

If Funcionarios_Ferias.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR MOVIMENTACAO NO BANCO RODBEL CASO NAO EXISTA NO CADASTRO
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionarios_Ferias.RecordCount
   
   Funcionarios_Ferias.MoveFirst
   
   While Not Funcionarios_Ferias.EOF
'         If Val(Funcionarios_Ferias!CHAPA) = 1429 Then
'             nGrvOk = 0
'         End If
         
         nGrvOk = 0
         
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE RELACIONAMENTO CRED_PESSOAS e PESSOAS,
         Rem CASO EXISTA, VERIFICAR SE JA FOI DEMITIDO(3) OU JA ESTEJA AFASTADO(4) CASO SIM IGNORAR REGISTRO
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Ferias!CHAPA
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)
         If rRs.RecordCount = 0 Then
            nsStatusMsg = 0
         Else
            Rem VERIFICAR O INTERVALO DO PERIODO DE FERIAS E TAMBEM
            Rem O sStatusMsg DO FUNCIONARIO, CASO SATISFACA, ENTRA NO AFASTAMENTO
            Rem NO CASO ENTRANDO DE FERIAS sStatusMsg = 10
            nsStatusMsg = 0
            
            If Funcionarios_Ferias!INICPROGFERIAS1 <= CDate(c_Data_Ferias) And _
               Funcionarios_Ferias!FIMPROGFERIAS1 >= CDate(c_Data_Ferias) Then 'CASO A DATA DE FERIAS ESTEJA NO INTERVALO ENCONTRADO, ENTAO ESTA DE FERIAS
               If rRs!pessit_numero = 10 Then
                  nsStatusMsg = 0
               Else
                If rRs!pessit_numero = 8 Then 'SO PODERA ENTRAR DE FERIAS CASO ESTA DESBLOQUEADO, CASO ESTEJA COM BLOQUEO DE OUTRA ACAO, NADA FARA.
                   nsStatusMsg = 10
                End If
               End If
            End If
            
            Rem VERIFICAR O INTERVALO DO PERIODO DE FERIAS E TAMBEM
            Rem O sStatusMsg DO FUNCIONARIO, NO CASO VOLTANDO DE FERIAS sStatusMsg = 2
            If Funcionarios_Ferias!FIMPROGFERIAS1 < CDate(c_Data_Ferias) Then
               If rRs!pessit_numero = 10 Then
                  nsStatusMsg = 8
               End If
            End If
            
         End If
         
         Rem NO CASO O sStatusMsg PARA ATUALIZAR OS DADOS SERA 8 E 10 RESPECTIVAMENTE
         If nsStatusMsg = 8 Or nsStatusMsg = 10 Then
            
            Rem ************************************************************************************************
            Rem ********* GRAVAR AS TABELAS DE PESSOAS, CREDENCIAIS********************************************
            Rem ************************************************************************************************
            
            Rem *************  PESSOAS  ***************
            Rem  ATUALIZACAO DO CAMPO PESSIST_NUMERO sStatusMsg = 3
            Set cFields = New Collection
            cFields.Add Funcionarios_Ferias!CHAPA
            cFields.Add "" 'NOME
            cFields.Add "" 'CPF
            cFields.Add "" 'OBSERVACAO
            cFields.Add "" 'E-MAIL
            cFields.Add "" 'IDENTIDADE
            cFields.Add "" 'CENTRO DE CUSTO
            cFields.Add "" 'FLAGREENTRADA
            cFields.Add "" 'MASTER
            cFields.Add "" 'FOTO
            cFields.Add "" 'SENHA
            cFields.Add "" 'RAMAL
            cFields.Add "" 'CAMPOPERSO1
            cFields.Add "" 'CAMPOPERSO2
            cFields.Add "" 'CAMPOPERSO3
            cFields.Add "" 'CAMPOPERSO4
            cFields.Add "" 'CAMPOPERSO5
            cFields.Add nsStatusMsg 'PESSIT_NUMERO
            If nGrvOk = 2 Then
               cFields.Add "" 'DTSITINICIAL
               cFields.Add "" 'DTSITFINAL
            Else
               cFields.Add Funcionarios_Ferias!INICPROGFERIAS1 'DTSITINICIAL
               cFields.Add Funcionarios_Ferias!FIMPROGFERIAS1 'DTSITFINAL
            End If
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'TOK_NUMEROSERIE
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 13)
            
            Rem ************************************************************************************************
            
            Rem *************  CREDENCIAIS  ***************
            rRs.MoveFirst
            While Not rRs.EOF
                Set cFields = New Collection
                cFields.Add rRs!CRED_NUMERO 'CRED_NUMERO
                cFields.Add "" 'CRED_FLAGREENTRADA
                cFields.Add "" 'CRED_AREANOMOMENTO
                cFields.Add "" 'CRED_DTVALIDINICIAL
                cFields.Add "" 'CRED_DTVALIDFINAL
                If nsStatusMsg = 8 Then
                   cFields.Add "0" 'CRED_BLOQUEADA
                Else
                   cFields.Add "1" 'CRED_BLOQUEADA
                End If
                cFields.Add "" 'CRED_FLAGCATRACA
                cFields.Add "" 'CRTI_NUMERO
                cFields.Add "" 'CRED_MASTER
                cFields.Add "" 'CRED_ULTPASSAGEM
                cFields.Add "" 'CRED_EQPINUMERO
                cFields.Add "" 'CRED_SENHA
                cFields.Add "" 'CRED_PROV
                cFields.Add "" 'CRED_PROVDEVOLV
                cFields.Add "" 'CRED_NUMRA
                
                Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 11)
                rRs.MoveNext
            Wend
            Rem ************************************************************************************************
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Ferias"
            sCodFun = Funcionarios_Ferias!CHAPA
            If nsStatusMsg = 8 Then
               sMsg = "Retorno de Ferias Periodo de " & Funcionarios_Ferias!INICPROGFERIAS1 & " a " & Funcionarios_Ferias!FIMPROGFERIAS1
            Else
               sMsg = "Entrando de Ferias Periodo de " & Funcionarios_Ferias!INICPROGFERIAS1 & " a " & Funcionarios_Ferias!FIMPROGFERIAS1
            End If
            
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            nGravou = nGravou + 1
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionarios_Ferias.EOF Then
            Funcionarios_Ferias.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If


ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Ferias"
   sCodFun = "0000"
   sMsg = "Nao ha Ferias no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing



Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Ferias"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

'ADOConnectRm.RollbackTrans
'ADOConnectRodbel.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DA INCLUIR SECOES
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionario_inclui_Secao(ByVal sBdRm As String, _
                                         ByVal sBdRodBel As String, _
                                         ByVal c_Data_Hoje As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

cFields.Add "'" & c_Data_Hoje & "'", "c_Data_Hoje"

Set Funcionario_inclui_Secao = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 20)

nGravou = 0
nLidos = 0

If Funcionario_inclui_Secao.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR MOVIMENTACAO NO BANCO RODBEL CASO EXISTA NO CADASTRO
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionario_inclui_Secao.RecordCount

   Funcionario_inclui_Secao.MoveFirst
   
   While Not Funcionario_inclui_Secao.EOF
'         If Val(Funcionario_inclui_Secao!CHAPA) = 1429 Then
'             nGrvOk = 0
'         End If
         
         nGrvOk = 0
         
         
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE ESTRUTURA OUVE MUDANCA DE SECAO NESTE FUNCIONARIO
         Rem CASO TENHA, ATUALIZAR O CTAB. PESSOAS O CAMPO EST_NUMERO COM A NOVA SECAO
         
         Set cFields = New Collection
         cFields.Add Funcionario_inclui_Secao!CODCOLIGADA
         cFields.Add Replace(Funcionario_inclui_Secao!CODIGO, ".", "")
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 23)
         
         If rRs.RecordCount = 1 Then 'CASO ACHE ALGUM REGISTRO LER O PROXIMO
            GoTo PROXIMO_REGISTRO
         End If
         
         Set cFields = New Collection
         cFields.Add Funcionario_inclui_Secao!CHAPA
         
         
         If nGrvOk = 0 Then
            
            Rem ************************************************************************************************
            Rem ********* atualizar AS TABELA DE ESTRUTURA  ********************************************
            Rem ************************************************************************************************
            
            Rem *************  estrutura  ***************

            Set cFields = New Collection
            cFields.Add Funcionario_inclui_Secao!CODCOLIGADA
            cFields.Add Replace(Funcionario_inclui_Secao!CODSECAO, ".", "") 'CENTRO DE CUSTO
            cFields.Add Funcionario_inclui_Secao!DESCRICAO
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 22)
            
            nGravou = nGravou + 1
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "inclusão Setor"
            sCodFun = Funcionario_inclui_Secao!CODCOLIGADA
            sMsg = "Inclusao de Setor " & Str(Funcionario_inclui_Secao!CODSECAO) & "-" & Funcionario_inclui_Secao!DESCRICAO
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionario_inclui_Secao.EOF Then
            Funcionario_inclui_Secao.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Inclusao Setor"
   sCodFun = "0000"
   sMsg = "Nao houve inclusao de Secao no periodo de " & c_Data_Hoje
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Inclusao Setor"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DA MUDANÇA DE SEÇÃO
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionario_Historico_Secao(ByVal sBdRm As String, _
                                            ByVal sBdRodBel As String, _
                                            ByVal c_Data_Hoje As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

cFields.Add "'" & c_Data_Hoje & "'", "c_Data_Hoje"

Set Funcionario_Historico_Secao = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 14)

nGravou = 0
nLidos = 0

If Funcionario_Historico_Secao.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR MOVIMENTACAO NO BANCO RODBEL CASO EXISTA NO CADASTRO
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionario_Historico_Secao.RecordCount

   Funcionario_Historico_Secao.MoveFirst
   
   While Not Funcionario_Historico_Secao.EOF
'         If Val(Funcionario_Historico_Secao!CHAPA) = 1429 Then
'             nGrvOk = 0
'         End If
         
         nGrvOk = 0
         
         
         Rem CRITICA PARA SABER DA EXISTENCIA NA TAB. DE ESTRUTURA OUVE MUDANCA DE SECAO NESTE FUNCIONARIO
         Rem CASO TENHA, ATUALIZAR O CTAB. PESSOAS O CAMPO EST_NUMERO COM A NOVA SECAO
         
         Set cFields = New Collection
         cFields.Add Funcionario_Historico_Secao!CHAPA
         cFields.Add Replace(Funcionario_Historico_Secao!CODSECAO, ".", "")
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 15)
         
         If rRs.RecordCount = 0 Then 'CASO NAO ACHE NENHUM REGISTRO LER O PROXIMO
            GoTo PROXIMO_REGISTRO
         End If
         
         Set cFields = New Collection
         cFields.Add Funcionario_Historico_Secao!CHAPA
         
         Rem CRITICA NO CASO DA NÃO EXISTENCIA DO NO SETOR NO CADASTRO DA RODBEL
         If rRs!ENCONTRADO = 0 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Mudança Setor"
            sCodFun = Funcionario_Historico_Secao!CHAPA
            sMsg = "Secao do RM nao Cadastada na RodBel  - " & Funcionario_Historico_Secao!CODSECAO
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            nGrvOk = 1
         
         End If
         
         Rem NO CASO DA SECAO DO FUNCIONARIO SER A MESMA, LER OUTRO REGISTRO
         If rRs!ENCONTRADO = 0 Then
            GoTo PROXIMO_REGISTRO
         End If
         
         
         If nGrvOk = 0 Then
            
            Rem ************************************************************************************************
            Rem ********* atualizar AS TABELA DE Funcionario_Historico_Secao            ********************************************
            Rem ************************************************************************************************
            
            Rem *************  PESSOAS  ***************

            Set cFields = New Collection
            cFields.Add Funcionario_Historico_Secao!CHAPA
            cFields.Add "" 'NOME
            cFields.Add "" 'CPF
            cFields.Add "" 'OBSERVACAO
            cFields.Add "" 'E-MAIL
            cFields.Add "" 'IDENTIDADE
            cFields.Add Replace(Funcionario_Historico_Secao!CODSECAO, ".", "") 'CENTRO DE CUSTO
            cFields.Add "" 'FLAGREENTRADA
            cFields.Add "" 'MASTER
            cFields.Add "" 'FOTO
            cFields.Add "" 'SENHA
            cFields.Add "" 'RAMAL
            cFields.Add "" 'CAMPOPERSO1
            cFields.Add "" 'CAMPOPERSO2
            cFields.Add "" 'CAMPOPERSO3
            cFields.Add "" 'CAMPOPERSO4
            cFields.Add "" 'CAMPOPERSO5
            cFields.Add "" 'PESSIT_NUMERO
            cFields.Add "" 'DTSITINICIAL
            cFields.Add "" 'DTSITFINAL
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'TOK_NUMEROSERIE
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 16)
            
            nGravou = nGravou + 1
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Mudança Setor"
            sCodFun = Funcionario_Historico_Secao!CHAPA
            sMsg = "Mudanca de Setor para " & Str(Funcionario_Historico_Secao!CODSECAO)
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionario_Historico_Secao.EOF Then
            Funcionario_Historico_Secao.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Mudança Setor"
   sCodFun = "0000"
   sMsg = "Nao ha Mudancas de Secao no periodo de " & c_Data_Hoje
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing

Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Mudança Setor"
sCodFun = Funcionario_Historico_Secao!CHAPA
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DO ATUALIZACAO DOS FUNCIONARIOS AFASTADOS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionarios_Afastado(ByVal sBdRm As String, _
                                      ByVal sBdRodBel As String, _
                                      ByVal c_Data_Afastado As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset 'record set para auxilio de acesso ao banco
Dim dDataCritica As Date

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

cFields.Add c_Data_Afastado, "c_Data_Afastado"

Set Funcionarios_Afastado = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 17)

nGravou = 0
nLidos = 0

If Funcionarios_Afastado.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR TABELA PESSOAS NO BANCO RODBEL CASO SATISFACA CONDICOES
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionarios_Afastado.RecordCount
   
   Funcionarios_Afastado.MoveFirst
   
   While Not Funcionarios_Afastado.EOF
'         If Val(Funcionarios_Afastado!CHAPA) = 1429 Then
'             nGrvOk = 0
'         End If
         
         nGrvOk = 0
         
         
         Rem SELECIONAR O REGISTRO PESSOAS DA RODBEL PARA VERIFICAR O STATUS DO CAMPO PESSIT_NUMERO E
         Rem CRED_BLOQUEADA DA TABELA CREDENCIAIS, E TESTAR A SITUACAO QUE VEIO DA RM
         Set cFields = New Collection
         cFields.Add Funcionarios_Afastado!CHAPA
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)
         
         If rRs.RecordCount = 0 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Afastados"
            sCodFun = Funcionarios_Afastado!CHAPA
            sMsg = "Tentativa de Afastar um funcionario nao Cadastrado na RodBel."
            
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            GoTo PROXIMO_REGISTRO
         Else
            nsStatusMsg = 0
            Rem VERIFICAR O A DATA FINAL :
            Rem CASO SEJA NULL OU > QUA A DATA DE HOJE, ESTARA AFASTADO
            Rem NO CASO DA DATA FINAL < DATA DE HOJE, OUVE RETORNO
            Rem NO RESULTADO ACIMA COMPARAR O CAMPO PESSIT_NUMERO E ATUALIZAR COMFORME A CRITICA ABAIXO
            dDataCritica = CDate(Trim(Mid$(c_Data_Afastado, 8, 2) & "/" & Mid$(c_Data_Afastado, 6, 2) & "/" & Mid$(c_Data_Afastado, 2, 4)))
            
            If IsNull(Funcionarios_Afastado!DTFINAL) Or _
               Funcionarios_Afastado!DTFINAL > dDataCritica Then
               If rRs!pessit_numero = 12 Then 'CASO JA ESTEJA COM A OPCAO 5 , JA ESTA BLOQUEADO
                  nsStatusMsg = 0
               Else
                  If rRs!pessit_numero = 8 Then 'CASO ESTEJA LIBERADO, SERA BLOQUEADO COMA OPCAO 12
                     nsStatusMsg = 12
                  End If
               End If
            ElseIf Funcionarios_Afastado!DTFINAL < dDataCritica Then
               If rRs!pessit_numero = 8 Then
                  nsStatusMsg = 0
               Else
                  If rRs!pessit_numero = 12 Then 'CASO ESTEJA EM BLOQUEIO DE AFASTAMENTO, SERA LIBERADO COM A OPCAO 8
                     nsStatusMsg = 8
                  End If
               End If
            End If
         
         End If
         
         Rem NO CASO O sStatusMsg PARA ATUALIZAR OS DADOS SERRA 2 E 5 RESPECTIVAMENTE
         If nsStatusMsg = 8 Or nsStatusMsg = 12 Then
            Rem ************************************************************************************************
            Rem ********* GRAVAR AS TABELAS DE PESSOAS, CREDENCIAIS********************************************
            Rem ************************************************************************************************
            
            Rem *************  PESSOAS  ***************
            Rem  ATUALIZACAO DO CAMPO PESSIST_NUMERO sStatusMsg = 5
            Set cFields = New Collection
            cFields.Add Funcionarios_Afastado!CHAPA
            cFields.Add "" 'NOME
            cFields.Add "" 'CPF
            cFields.Add "" 'OBSERVACAO
            cFields.Add "" 'E-MAIL
            cFields.Add "" 'IDENTIDADE
            cFields.Add "" 'CENTRO DE CUSTO
            cFields.Add "" 'FLAGREENTRADA
            cFields.Add "" 'MASTER
            cFields.Add "" 'FOTO
            cFields.Add "" 'SENHA
            cFields.Add "" 'RAMAL
            cFields.Add "" 'CAMPOPERSO1
            cFields.Add "" 'CAMPOPERSO2
            cFields.Add "" 'CAMPOPERSO3
            cFields.Add "" 'CAMPOPERSO4
            cFields.Add "" 'CAMPOPERSO5
            cFields.Add nsStatusMsg 'PESSIT_NUMERO
            cFields.Add "" 'DTSITINICIAL
            cFields.Add "" 'DTSITFINAL
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'TOK_NUMEROSERIE
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 10)
                
            Rem ************************************************************************************************
            
            rRs.MoveFirst
            While Not rRs.EOF
                
                Rem *************  CREDENCIAIS  ***************
                Set cFields = New Collection
                cFields.Add rRs!CRED_NUMERO 'CRED_NUMERO
                cFields.Add "" 'CRED_FLAGREENTRADA
                cFields.Add "" 'CRED_AREANOMOMENTO
                cFields.Add "" 'CRED_DTVALIDINICIAL
                cFields.Add "" 'CRED_DTVALIDFINAL
                If nsStatusMsg = 8 Then ' LIBERADA
                   cFields.Add "0" 'CRED_BLOQUEADA
                Else
                   cFields.Add "1" 'CRED_BLOQUEADA
                End If
                cFields.Add "" 'CRED_FLAGCATRACA
                cFields.Add "" 'CRTI_NUMERO
                cFields.Add "" 'CRED_MASTER
                cFields.Add "" 'CRED_ULTPASSAGEM
                cFields.Add "" 'CRED_EQPINUMERO
                cFields.Add "" 'CRED_SENHA
                cFields.Add "" 'CRED_PROV
                cFields.Add "" 'CRED_PROVDEVOLV
                cFields.Add "" 'CRED_NUMRA
                
                Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 11)
                rRs.MoveNext
            Wend
            Rem ************************************************************************************************
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Afastados"
            sCodFun = Funcionarios_Afastado!CHAPA
            If nsStatusMsg = 8 Then
               sMsg = "Retorno de Afastamento Periodo de " & Funcionarios_Afastado!DTINICIO & " a " & Funcionarios_Afastado!DTFINAL
            Else
               sMsg = "Afastando no Inicio de " & Funcionarios_Afastado!DTINICIO & IIf(IsNull(Funcionarios_Afastado!DTFINAL), " ", " a " & Funcionarios_Afastado!DTFINAL)
            End If
            
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            nGravou = nGravou + 1
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionarios_Afastado.EOF Then
            Funcionarios_Afastado.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If


ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Afastados"
   sCodFun = "0000"
   sMsg = "Nao ha Afastamento no periodo de " & c_Data_Afastado
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing



Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Afastados"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

'ADOConnectRm.RollbackTrans
'ADOConnectRodbel.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   CAMADA DE NEGOCIOS DO ATUALIZACAO DOS FUNCIONARIOS DE ATESTADOS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Funcionarios_Atestado(ByVal sBdRm As String, _
                                      ByVal sBdRodBel As String, _
                                      ByVal c_Data_Atestado As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnectRm As ADODB.Connection
Dim ADOConnectRodbel As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nGravou As Double 'quantidade le registro atualizados no banco para saber depois se instancia o commit ou o rollback
Dim nsStatusMsg As Integer 'Flag para gravacao do registro
Dim nGrvOk As Integer 'verificar quantidade re registros gravados
Dim nLidos As Integer 'verificar a quantidade de registros lidos
Dim rRs As ADODB.Recordset 'record set para auxilio de acesso ao banco
Dim dDataCritica As Date

On Error GoTo Erro


Set ADOConnectRm = New ADODB.Connection
Set ADOConnectRodbel = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection

Set ADOConnectRm = cConect.Coneccao_OLEDB(sBdRm, "A")
Set ADOConnectRodbel = cConect.Coneccao_OLEDB(sBdRodBel, "A")

cFields.Add c_Data_Atestado, "c_Data_Atestado"

Set Funcionarios_Atestado = RegBanco_Critica_Pessoas(ADOConnectRm, cFields, 20)

nGravou = 0
nLidos = 0

If Funcionarios_Atestado.RecordCount > 0 Then ' CASO ACHE FUNCIONARIO, ATUALIZAR TABELA PESSOAS NO BANCO RODBEL CASO SATISFACA CONDICOES
   frmNotfIc!Pr_Prog.Visible = True
   frmNotfIc!Pr_Prog.Min = 0
   frmNotfIc!Pr_Prog.Max = Funcionarios_Atestado.RecordCount
   
   Funcionarios_Atestado.MoveFirst
   
   While Not Funcionarios_Atestado.EOF
'         If Val(Funcionarios_Atestado!CHAPA) = 1429 Then
'             nGrvOk = 0
'         End If
         
         nGrvOk = 0
         
         
         Set cFields = New Collection
         cFields.Add Funcionarios_Atestado!CHAPA
         Rem SELECIONAR O REGISTRO PESSOAS DA RODBEL PARA VERIFICAR O STATUS DO CAMPO PESSIT_NUMERO E
         Rem CRED_BLOQUEADA DA TABELA CREDENCIAIS, E TESTAR A SITUACAO QUE VEIO DA RM
         Set rRs = New ADODB.Recordset
         Set rRs = RegBanco_Critica_Rel(ADOConnectRodbel, cFields, 8)
         If rRs.RecordCount = 0 Then
            sStatusMsg = "1"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Atestados"
            sCodFun = Funcionarios_Atestado!CHAPA
            sMsg = "Tentativa de Afastar um funcionario nao cadastrado na RodBel."
            
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
            GoTo PROXIMO_REGISTRO
         Else
            nsStatusMsg = 0
            Rem VERIFICAR O A DATA FINAL :
            Rem CASO SEJA NULL OU > QUA A DATA DE HOJE, ESTARA Atestado
            Rem NO CASO DA DATA FINAL < DATA DE HOJE, OUVE RETORNO
            Rem NO RESULTADO ACIMA COMPARAR O CAMPO PESSIT_NUMERO E ATUALIZAR COMFORME A CRITICA ABAIXO
            dDataCritica = CDate(c_Data_Atestado)
            If IsNull(Funcionarios_Atestado!DTFINAL) Or _
               Funcionarios_Atestado!DTFINAL > dDataCritica Then
               If rRs!pessit_numero = 12 Then 'CASO JA ESTEJA COM A OPCAO 12 , JA ESTA BLOQUEADO
                  nsStatusMsg = 0
               Else
                  If rRs!pessit_numero = 8 Then 'CASO ESTEJA LIBERADO, SERA BLOQUADO COMA OPCAO 12
                     nsStatusMsg = 12
                  End If
               End If
            ElseIf Funcionarios_Atestado!DTFINAL < dDataCritica Then
               If rRs!pessit_numero = 8 Then
                  nsStatusMsg = 0
               Else
                  If rRs!pessit_numero = 12 Then 'CASO ESTEJA EM BLOQUEIO DE Atestado, SERA LIBERADO COM A OPCAO 2
                     nsStatusMsg = 8
                  End If
               End If
            End If
         
         End If
         
         Rem NO CASO O sStatusMsg PARA ATUALIZAR OS DADOS SERRA 2 E 4 RESPECTIVAMENTE
         If nsStatusMsg = 8 Or nsStatusMsg = 12 Then
            
            Rem ************************************************************************************************
            Rem ********* GRAVAR AS TABELAS DE PESSOAS, CREDENCIAIS********************************************
            Rem ************************************************************************************************
            
            Rem *************  PESSOAS  ***************
            Rem  ATUALIZACAO DO CAMPO PESSIST_NUMERO sStatusMsg = 6
            Set cFields = New Collection
            cFields.Add Funcionarios_Atestado!CHAPA
            cFields.Add "" 'NOME
            cFields.Add "" 'CPF
            cFields.Add "" 'OBSERVACAO
            cFields.Add "" 'E-MAIL
            cFields.Add "" 'IDENTIDADE
            cFields.Add "" 'CENTRO DE CUSTO
            cFields.Add "" 'FLAGREENTRADA
            cFields.Add "" 'MASTER
            cFields.Add "" 'FOTO
            cFields.Add "" 'SENHA
            cFields.Add "" 'RAMAL
            cFields.Add "" 'CAMPOPERSO1
            cFields.Add "" 'CAMPOPERSO2
            cFields.Add "" 'CAMPOPERSO3
            cFields.Add "" 'CAMPOPERSO4
            cFields.Add "" 'CAMPOPERSO5
            cFields.Add nsStatusMsg 'PESSIT_NUMERO
            cFields.Add "" 'DTSITINICIAL
            cFields.Add "" 'DTSITFINAL
            cFields.Add "" 'CRED_PROV
            cFields.Add "" 'TOK_NUMEROSERIE
            
            Call RegBanco_Critica_Pessoas(ADOConnectRodbel, cFields, 10)
            
            Rem ************************************************************************************************
            
            Rem *************  CREDENCIAIS  ***************
            rRs.MoveFirst
            While Not rRs.EOF
                Set cFields = New Collection
                cFields.Add Funcionarios_Atestado!CHAPA 'CRED_NUMERO
                cFields.Add "" 'CRED_FLAGREENTRADA
                cFields.Add "" 'CRED_AREANOMOMENTO
                cFields.Add "" 'CRED_DTVALIDINICIAL
                cFields.Add "" 'CRED_DTVALIDFINAL
                If nsStatusMsg = 8 Then
                   cFields.Add "0" 'CRED_BLOQUEADA
                Else
                   cFields.Add "1" 'CRED_BLOQUEADA
                End If
                cFields.Add "" 'CRED_FLAGCATRACA
                cFields.Add "" 'CRTI_NUMERO
                cFields.Add "" 'CRED_MASTER
                cFields.Add "" 'CRED_ULTPASSAGEM
                cFields.Add "" 'CRED_EQPINUMERO
                cFields.Add "" 'CRED_SENHA
                cFields.Add "" 'CRED_PROV
                cFields.Add "" 'CRED_PROVDEVOLV
                cFields.Add "" 'CRED_NUMRA
                
                Call RegBanco_Critica_Cred(ADOConnectRodbel, cFields, 11)
                rRs.MoveNext
            Wend
            Rem ************************************************************************************************
            
            sStatusMsg = "0"
            sData = Format(Now(), "dd/mm/yyyy")
            sHora = Format(Now(), "HH:MM")
            sTipo = "Atestados"
            sCodFun = Funcionarios_Atestado!CHAPA
            If nsStatusMsg = 8 Then
               sMsg = "Retorno de Atestado Periodo de " & Funcionarios_Atestado!DTINICIO & " a " & Funcionarios_Atestado!DTFINAL
            Else
               sMsg = "Atestado no Inicio de " & Funcionarios_Atestado!DTINICIO & IIf(IsNull(Funcionarios_Atestado!DTFINAL), " ", " a " & Funcionarios_Atestado!DTFINAL)
            End If
            
            Set cFields = New Collection
            cFields.Add sStatusMsg & ";" & _
                               sData & ";" & _
                               sHora & ";" & _
                               sCodFun & ";" & _
                               sTipo & ";" & _
                               sMsg
            
            Call Gerar_Situacao_Log(cFields)
         
            ADOConnectRodbel.CommitTrans
            ADOConnectRodbel.BeginTrans
            nGravou = nGravou + 1
            
         End If

PROXIMO_REGISTRO:
         
         If Not Funcionarios_Atestado.EOF Then
            Funcionarios_Atestado.MoveNext
            nLidos = nLidos + 1
            frmNotfIc!Pr_Prog.Value = nLidos
         End If
   Wend

End If


ADOConnectRm.CommitTrans
ADOConnectRodbel.CommitTrans
ADOConnectRm.Close
ADOConnectRodbel.Close

If nGravou = 0 Then
   sStatusMsg = "0"
   sData = Format(Now(), "dd/mm/yyyy")
   sHora = Format(Now(), "HH:MM")
   sTipo = "Atestados"
   sCodFun = "0000"
   sMsg = "Nao ha Atestado no periodo de " & Format(Now(), "dd/mm/yyyy hh:mm")
   Set cFields = New Collection
   cFields.Add sStatusMsg & ";" & _
                      sData & ";" & _
                      sHora & ";" & _
                      sCodFun & ";" & _
                      sTipo & ";" & _
                      sMsg
   
   Call Gerar_Situacao_Log(cFields)
End If

Set ADOConnectRm = Nothing
Set ADOConnectRodbel = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRs = Nothing



Exit Function

Erro:

If Err.Number = 50000 Then
   sStatusMsg = "1"
Else
   sStatusMsg = "2"
End If
sData = Format(Now(), "dd/mm/yyyy")
sHora = Format(Now(), "HH:MM")
sTipo = "Atestados"
sCodFun = "0000"
sMsg = Err.Description
Set cFields = New Collection
cFields.Add sStatusMsg & ";" & sData & ";" & sHora & ";" & sCodFun & ";" & sTipo & ";" & sMsg
Call Gerar_Situacao_Log(cFields)
Err.Clear

GoTo PROXIMO_REGISTRO

'ADOConnectRm.RollbackTrans
'ADOConnectRodbel.RollbackTrans

'Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA CRITICA E PREPARA OS DADOS PARA MANU
'   TENCAO NO BANCO, EM ESPECIAL A TABELA DE PESSOAS NO RODBEL
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function RegBanco_Critica_Pessoas(ByVal cCONECCAO As Connection, _
                                         ByVal cFields As Collection, _
                                         ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro

Dim nx As Integer
Dim ssql As String
Dim sName As String


If nAction > 1 And nAction <> 9 And nAction <> 12 And nAction < 14 Then
   'Critica dos Campos dos Vendedor

    If nAction = 3 And Len(Trim(cFields(1))) = 0 Then
       Err.Raise 50000, Err.Source, "Codigo do funcionario Invalido"
    End If
    If nAction = 3 And Len(Trim(cFields(2))) = 0 Then
       Err.Raise 50000, Err.Source, "Nome do funcionario invalido"
    End If
    If nAction = 3 And Len(Trim(cFields(3))) = 0 Then
       Err.Raise 50000, Err.Source, "CPF do funcionario invalido"
    End If
    If nAction = 3 And Len(Trim(cFields(6))) = 0 Then
       Err.Raise 50000, Err.Source, "carteira de identidade do funcionario invalido"
    End If
    If nAction = 3 And Len(Trim(cFields(7))) = 0 Then
       Err.Raise 50000, Err.Source, "Centro de Custo do funcionario invalido"
    End If
   'Por aspas,Nulls,virguas nos campos correspondentes
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 2 To 22
                 If Len(cFields(nx)) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "NULL", , , nx - 1
                 End If

                'Verifica campos numéricos NULL
                'Case 32, 35
                '    If Val(cFields(nx)) = 0 Then
                '        'sName = cFields(nx).Name
                '        cFields.Remove nx
                '        cFields.Add "NULL", , , nx - 1
                '    End If
            End Select
        End If

        If cFields(nx) <> "NULL" Then
            Select Case nx
                'Altera campos alfanuméricos
                Case 2, 3, 4, 5, 6, 10 To 17, 19, 20, 21, 22
                    ssql = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1

                ' Altera campos tipo data
                Case 0
                    ssql = "#" & Format(cFields(nx), "mm/dd/yyyy") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1
                 'Altera campos tipo data da alteração
                Case 0
                    ssql = "#" & Format(cFields(nx), "mm/dd/yyyy HH:MM:SS") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1
                'Altera campos Boolean
                Case 0
                If Len(Trim(cFields(nx))) = 0 Then
                   If cFields(nx) Then
                       ssql = "1"
                   Else
                       ssql = "0"
                   End If
                   'sName = cFields(nx).Name
                   cFields.Remove nx
                   cFields.Add ssql, , , nx - 1
                End If
                'Altera campos numéricos fracionários
                Case 0
'                    If InStr(cFields(nx), ",") > 0 Then
                        ssql = Replace(cFields(nx), ".", "")
                        ssql = Replace(ssql, ",", ".")
                        cFields.Remove nx
                        cFields.Add ssql, , , nx - 1
'                    End If
            End Select
        End If

    Next nx
End If

Set RegBanco_Critica_Pessoas = RegBanco_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA CRITICA E PREPARA OS DADOS PARA MANU-
'   TENCAO NO BANCO, EM ESPECIAL A TABELA DE RELACIONAMENTO
'   CRED_PESSOAS NO RODBEL
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function RegBanco_Critica_Rel(ByVal cCONECCAO As Connection, _
                                     ByVal cFields As Collection, _
                                     ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim ssql As String
Dim sName As String

If nAction > 1 And nAction < 11 Then
   'Critica dos Campos dos Vendedor
'    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
'       Err.Raise 50000, Err.Source, "Data inicial Inválido"
'    End If
'    If Len(Trim(cFields(2))) = 0 Then
'       Err.Raise 50000, Err.Source, "dataFechamento Inválido"
'    End If
'    If Len(Trim(cFields(4))) = 0 Then
'       Err.Raise 50000, Err.Source, "CPF / CGC do VENDEDOR não Digitado"
'    End If
   
   'Por aspas,Nulls,virguas nos campos correspondentes
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 2 To 4
                 If Len(cFields(nx)) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "NULL", , , nx - 1
                 End If

                'Verifica campos numéricos NULL
                'Case 32, 35
                '    If Val(cFields(nx)) = 0 Then
                '        'sName = cFields(nx).Name
                '        cFields.Remove nx
                '        cFields.Add "NULL", , , nx - 1
                '    End If
            End Select
        End If

        If cFields(nx) <> "NULL" Then
            Select Case nx
                'Altera campos alfanuméricos
                Case 3, 4
                    ssql = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1

                ' Altera campos tipo data
                Case 0
                    ssql = "#" & Format(cFields(nx), "mm/dd/yyyy") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1
                 'Altera campos tipo data da alteração
                Case 0
                    ssql = "#" & Format(cFields(nx), "mm/dd/yyyy HH:MM:SS") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1
                'Altera campos Boolean
                Case 0
                If Len(Trim(cFields(nx))) = 0 Then
                   If cFields(nx) Then
                       ssql = "1"
                   Else
                       ssql = "0"
                   End If
                   'sName = cFields(nx).Name
                   cFields.Remove nx
                   cFields.Add ssql, , , nx - 1
                End If
                'Altera campos numéricos fracionários
                Case 0
'                    If InStr(cFields(nx), ",") > 0 Then
                        ssql = Replace(cFields(nx), ".", "")
                        ssql = Replace(ssql, ",", ".")
                        cFields.Remove nx
                        cFields.Add ssql, , , nx - 1
'                    End If
            End Select
        End If

    Next nx
End If

Set RegBanco_Critica_Rel = RegBanco_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function RegBanco_Critica_Estrutura(ByVal cCONECCAO As Connection, _
                                           ByVal cFields As Collection, _
                                           ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim ssql As String
Dim sName As String

If nAction > 1 And nAction < 11 Then
   'Critica dos Campos dos Vendedor
'    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
'       Err.Raise 50000, Err.Source, "Data inicial Inválido"
'    End If
'    If Len(Trim(cFields(2))) = 0 Then
'       Err.Raise 50000, Err.Source, "dataFechamento Inválido"
'    End If
'    If Len(Trim(cFields(4))) = 0 Then
'       Err.Raise 50000, Err.Source, "CPF / CGC do VENDEDOR não Digitado"
'    End If
   
   'Por aspas,Nulls,virguas nos campos correspondentes
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 2 To 15
                 If Len(cFields(nx)) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "NULL", , , nx - 1
                 End If

                'Verifica campos numéricos NULL
                'Case 32, 35
                '    If Val(cFields(nx)) = 0 Then
                '        'sName = cFields(nx).Name
                '        cFields.Remove nx
                '        cFields.Add "NULL", , , nx - 1
                '    End If
            End Select
        End If

        If cFields(nx) <> "NULL" Then
            Select Case nx
                'Altera campos alfanuméricos
                Case 0
                    ssql = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1

                ' Altera campos tipo data
                Case 0
                    ssql = "#" & Format(cFields(nx), "mm/dd/yyyy") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1
                 'Altera campos tipo data da alteração
                Case 0
                    ssql = "#" & Format(cFields(nx), "mm/dd/yyyy HH:MM:SS") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1
                'Altera campos Boolean
                Case 0
                If Len(Trim(cFields(nx))) = 0 Then
                   If cFields(nx) Then
                       ssql = "1"
                   Else
                       ssql = "0"
                   End If
                   'sName = cFields(nx).Name
                   cFields.Remove nx
                   cFields.Add ssql, , , nx - 1
                End If
                'Altera campos numéricos fracionários
                Case 0
'                    If InStr(cFields(nx), ",") > 0 Then
                        ssql = Replace(cFields(nx), ".", "")
                        ssql = Replace(ssql, ",", ".")
                        cFields.Remove nx
                        cFields.Add ssql, , , nx - 1
'                    End If
            End Select
        End If

    Next nx
End If

Set RegBanco_Critica_Estrutura = RegBanco_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description

End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA CRITICA E PREPARA OS DADOS PARA MANU
'   TENCAO NO BANCO, EM ESPECIAL A TABELA DE CREDENCIAIS  NO
'   RODBEL
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function RegBanco_Critica_Cred(ByVal cCONECCAO As Connection, _
                                      ByVal cFields As Collection, _
                                      ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim ssql As String
Dim sName As String

If nAction > 1 And nAction <= 11 Then
   'Critica dos Campos dos Vendedor
'    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
'       Err.Raise 50000, Err.Source, "Data inicial Inválido"
'    End If
'    If Len(Trim(cFields(2))) = 0 Then
'       Err.Raise 50000, Err.Source, "dataFechamento Inválido"
'    End If
'    If Len(Trim(cFields(4))) = 0 Then
'       Err.Raise 50000, Err.Source, "CPF / CGC do VENDEDOR não Digitado"
'    End If
   
   'Por aspas,Nulls,virguas nos campos correspondentes
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 2 To 11, 13, 14, 15
                 If Len(cFields(nx)) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "NULL", , , nx - 1
                 End If

                'Verifica campos numéricos NULL
                'Case 32, 35
                '    If Val(cFields(nx)) = 0 Then
                '        'sName = cFields(nx).Name
                '        cFields.Remove nx
                '        cFields.Add "NULL", , , nx - 1
                '    End If
            End Select
        End If

        If cFields(nx) <> "NULL" Then
            Select Case nx
                'Altera campos alfanuméricos
                Case 12
                    ssql = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1

                ' Altera campos tipo data
                Case 0
                    ssql = "#" & Format(cFields(nx), "mm/dd/yyyy") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1
                 'Altera campos tipo data da alteração
                Case 0
                    ssql = "#" & Format(cFields(nx), "mm/dd/yyyy HH:MM:SS") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1
                'Altera campos Boolean
                Case 0
                If Len(Trim(cFields(nx))) = 0 Then
                   If cFields(nx) Then
                       ssql = "1"
                   Else
                       ssql = "0"
                   End If
                   'sName = cFields(nx).Name
                   cFields.Remove nx
                   cFields.Add ssql, , , nx - 1
                End If
                'Altera campos numéricos fracionários
                Case 0
'                    If InStr(cFields(nx), ",") > 0 Then
                        ssql = Replace(cFields(nx), ".", "")
                        ssql = Replace(ssql, ",", ".")
                        cFields.Remove nx
                        cFields.Add ssql, , , nx - 1
'                    End If
            End Select
        End If

    Next nx
End If

Set RegBanco_Critica_Cred = RegBanco_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA CRITICA E PREPARA OS DADOS PARA MANU
'   TENCAO NO BANCO, EM ESPECIAL A TABELA DE CRED_ACESSO   NO
'   RODBEL
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function RegBanco_Critica_Cred_Acesso(ByVal cCONECCAO As Connection, _
                                             ByVal cFields As Collection, _
                                             ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim ssql As String
Dim sName As String

If nAction = 19 Then
   'Critica dos Campos dos Vendedor
'    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
'       Err.Raise 50000, Err.Source, "Data inicial Inválido"
'    End If
'    If Len(Trim(cFields(2))) = 0 Then
'       Err.Raise 50000, Err.Source, "dataFechamento Inválido"
'    End If
'    If Len(Trim(cFields(4))) = 0 Then
'       Err.Raise 50000, Err.Source, "CPF / CGC do VENDEDOR não Digitado"
'    End If
   
   'Por aspas,Nulls,virguas nos campos correspondentes
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 5 To 9
                 If Len(cFields(nx)) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "NULL", , , nx - 1
                 End If

                'Verifica campos numéricos NULL
                'Case 32, 35
                '    If Val(cFields(nx)) = 0 Then
                '        'sName = cFields(nx).Name
                '        cFields.Remove nx
                '        cFields.Add "NULL", , , nx - 1
                '    End If
            End Select
        End If

        If cFields(nx) <> "NULL" Then
            Select Case nx
                'Altera campos alfanuméricos
                Case 12
                    ssql = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1

                ' Altera campos tipo data
                Case 0
                    ssql = "#" & Format(cFields(nx), "mm/dd/yyyy") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1
                 'Altera campos tipo data da alteração
                Case 0
                    ssql = "#" & Format(cFields(nx), "mm/dd/yyyy HH:MM:SS") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add ssql, , , nx - 1
                'Altera campos Boolean
                Case 0
                If Len(Trim(cFields(nx))) = 0 Then
                   If cFields(nx) Then
                       ssql = "1"
                   Else
                       ssql = "0"
                   End If
                   'sName = cFields(nx).Name
                   cFields.Remove nx
                   cFields.Add ssql, , , nx - 1
                End If
                'Altera campos numéricos fracionários
                Case 0
'                    If InStr(cFields(nx), ",") > 0 Then
                        ssql = Replace(cFields(nx), ".", "")
                        ssql = Replace(ssql, ",", ".")
                        cFields.Remove nx
                        cFields.Add ssql, , , nx - 1
'                    End If
            End Select
        End If

    Next nx
End If

Set RegBanco_Critica_Cred_Acesso = RegBanco_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE DADOS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function RegBanco_dados(ByVal cCONECCAO As Connection, _
                               ByVal cFields As Collection, _
                               ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim rs As ADODB.Recordset
Dim ssql As String
Dim sDate As String
Dim nseq As String



Select Case nAction
'                  "PFUNC.INTEGRGERENCIAL, "
    Case 1 'Consultar CADASTRO DE PESSOAL VERIFICANDO A DATA DE ADMISSAO
           ssql = "SELECT " & _
                  "PFUNC.CHAPA, " & _
                  "PFUNC.NOME, " & _
                  "PFUNC.CODSECAO AS INTEGRGERENCIAL, " & _
                  "PPESSOA.CPF, " & _
                  "PPESSOA.CARTIDENTIDADE, " & _
                  "PFUNC.CODFUNCAO, " & _
                  "PFUNCAO.CARGO "
           
           ssql = ssql & _
                  "FROM PFUNC " & _
                  "INNER JOIN PPESSOA ON PFUNC.CODPESSOA = PPESSOA.CODIGO " & _
                  "INNER JOIN PFUNCAO ON PFUNC.CODFUNCAO = PFUNCAO.CODIGO AND PFUNC.CODCOLIGADA = PFUNCAO.CODCOLIGADA"
       
           ssql = ssql & _
                  " WHERE PFUNC.CODCOLIGADA = 1" & _
                  " AND PFUNC.DATAADMISSAO >= '" & cFields(1) & "'"
         
'           ssql = ssql & _
'                  " WHERE PFUNC.CODCOLIGADA = 1" & _
'                  " AND PFUNC.CHAPA = '0361'"
         
    Case 2 'Consultar CADASTRO DA RODBEL VERIFICANDO SE HA O REGISTRO DO FUNCIONARIO
           ssql = "SELECT " & _
                  "PES_NUMERO AS ENCONTRADO "
                  
           ssql = ssql & _
                  "FROM PESSOAS "
       
           ssql = ssql & _
                  " WHERE PES_NUMERO = " & cFields(1)
         
    Case 3 'INSERIR NA TABELA DE PESSOAS DA RODBEL OS NOVOS FUNCIONARIOS ADMITIDOS PELO SISTEMA RM
    
            ssql = "INSERT INTO PESSOAS (" & _
                   "PES_NUMERO, " & _
                   "PES_NOME, " & _
                   "PES_CPF, " & _
                   "PES_OBS, " & _
                   "PES_EMAIL, " & _
                   "PES_RG, " & _
                   "EST_NUMERO, " & _
                   "PES_FLAGREENTRADA, " & _
                   "PES_MASTER, " & _
                   "PES_FOTO, " & _
                   "PES_SENHA, " & _
                   "PES_RAMAL, " & _
                   "PES_CAMPOPERSO1, " & _
                   "PES_CAMPOPERSO2, " & _
                   "PES_CAMPOPERSO3, " & _
                   "PES_CAMPOPERSO4, " & _
                   "PES_CAMPOPERSO5, " & _
                   "PESSIT_NUMERO, " & _
                   "PES_DTSITINICIAL, " & _
                   "PES_DTSITFINAL, " & _
                   "CRED_PROV, " & _
                   "TOK_NUMEROSERIE "
        
            ssql = ssql & ") VALUES (" & _
                            cFields(1) & ", " & _
                            cFields(2) & ", " & _
                            cFields(3) & ", " & _
                            cFields(4) & ", " & _
                            cFields(5) & ", " & _
                            cFields(6) & ", " & _
                            cFields(7) & ", " & _
                            cFields(8) & ", " & _
                            cFields(9) & ", " & _
                            cFields(10) & ", " & _
                            cFields(11) & ", " & _
                            cFields(12) & ", " & _
                            cFields(13) & ", " & _
                            cFields(14) & ", " & _
                            cFields(15) & ", " & _
                            cFields(16) & ", " & _
                            cFields(17) & ", " & _
                            cFields(18) & ", " & _
                            cFields(19) & ", " & _
                            cFields(20) & ", " & _
                            cFields(21) & ", " & _
                            cFields(22) & ")"
                    
    Case 4 'INSERIR NA TABELA DE CREDENCIAIS DA RODBEL OS NOVOS FUNCIONARIOS ADMITIDOS PELO SISTEMA RM
    
           
           ssql = "INSERT INTO CREDENCIAIS (" & _
                  "CRED_NUMERO, " & _
                  "CRED_FLAGREENTRADA, " & _
                  "CRED_AREANOMOMENTO, " & _
                  "CRED_DTVALIDINICIAL, " & _
                  "CRED_DTVALIDFINAL, " & _
                  "CRED_BLOQUEADA, " & _
                  "CRED_FLAGCATRACA, " & _
                  "CRTI_NUMERO, " & _
                  "CRED_MASTER, " & _
                  "CRED_ULTPASSAGEM, " & _
                  "CRED_EQPINUMERO, " & _
                  "CRED_SENHA, " & _
                  "CRED_PROV, " & _
                  "CRED_PROVDEVOLV, " & _
                  "CRED_NUMRA "
           
           ssql = ssql & ") VALUES (" & _
                           cFields(1) & ", " & _
                           cFields(2) & ", " & _
                           cFields(3) & ", " & _
                           cFields(4) & ", " & _
                           cFields(5) & ", " & _
                           cFields(7) & ", " & _
                           cFields(6) & ", " & _
                           cFields(8) & ", " & _
                           cFields(9) & ", " & _
                           cFields(10) & ", " & _
                           cFields(11) & ", " & _
                           cFields(12) & ", " & _
                           cFields(13) & ", " & _
                           cFields(14) & ", " & _
                           cFields(15) & " )"

    Case 5 'INSERIR NA TABELA DE CRED_PESSOAS DA RODBEL OS NOVOS FUNCIONARIOS ADMITIDOS PELO SISTEMA RM
    
           ssql = "INSERT INTO CRED_PESSOAS (" & _
                  "CRED_NUMERO, " & _
                  "PES_NUMERO, " & _
                  "CRPES_DATAENTRADA, " & _
                  "CRPES_DATASAIDA"
                  
           ssql = ssql & ") VALUES (" & _
                           cFields(1) & ", " & _
                           cFields(2) & ", " & _
                           cFields(3) & ", " & _
                           cFields(4) & " )"

    Case 6 'Consultar TABELA DE ESTRUTURA DA RODBEL VERIFICANDO SE HA O REGISTRO DO CENTRO DE CUSTO CADASTRADO
           ssql = "SELECT " & _
                  "COUNT(*) AS ENCONTRADO "
                  
           ssql = ssql & _
                  "FROM ESTRUTURA "
       
           ssql = ssql & _
                  " WHERE EST_NUMERO = " & cFields(1)
    
    Case 7 'Consultar TABELA DE CRED_PESSOAS DA RODBEL VERIFICANDO SE HA O REGISTRO CADASTRADO
           ssql = "SELECT " & _
                  "PES_NUMERO AS ENCONTRADO "
                  
           ssql = ssql & _
                  "FROM CRED_PESSOAS "
       
           ssql = ssql & _
                  " WHERE PES_NUMERO = " & cFields(1)
    
    Case 8 'Consultar TABELA DE RELACIONAMENTO DE CRED_PESSOAS E PESSOAS DA RODBEL VERIFICANDO SE HA O REGISTRO CADASTRADO
           ssql = "SELECT "
           
           ssql = ssql & _
                  "(SELECT COUNT(*) " & _
                  "FROM CRED_PESSOAS " & _
                  "INNER JOIN CREDENCIAIS " & _
                  "ON CREDENCIAIS.CRED_NUMERO = CRED_PESSOAS.CRED_NUMERO " & _
                  "INNER JOIN PESSOAS " & _
                  "ON CRED_PESSOAS.PES_NUMERO = PESSOAS.PES_NUMERO " & _
                  "WHERE PESSOAS.PES_NUMERO = " & cFields(1) & " )" & _
                  "AS ENCONTRADO,"
           
           ssql = ssql & _
                  "CREDENCIAIS.CRED_BLOQUEADA, " & _
                  "PESSOAS.PESSIT_NUMERO, " & _
                  "CREDENCIAIS.CRED_NUMERO "
                  
           ssql = ssql & _
                  "FROM CRED_PESSOAS " & _
                  "INNER JOIN CREDENCIAIS " & _
                  "ON CREDENCIAIS.CRED_NUMERO = CRED_PESSOAS.CRED_NUMERO " & _
                  "INNER JOIN PESSOAS " & _
                  "ON CRED_PESSOAS.PES_NUMERO = PESSOAS.PES_NUMERO "
           
           ssql = ssql & _
                  " WHERE PESSOAS.PES_NUMERO = " & cFields(1)

           ssql = ssql & _
                  " GROUP BY CREDENCIAIS.CRED_NUMERO,CREDENCIAIS.CRED_BLOQUEADA, PESSOAS.PESSIT_NUMERO"
    
    Case 9 'Consultar CADASTRO DE PESSOAL VERIFICANDO A DATA DE DEMISSAO
           ssql = "SELECT " & _
                  "PFUNC.CHAPA, " & _
                  "PFUNC.NOME, " & _
                  "PFUNC.INTEGRGERENCIAL, " & _
                  "PPESSOA.CPF, " & _
                  "PPESSOA.CARTIDENTIDADE "
           
           ssql = ssql & _
                  "FROM PFUNC " & _
                  "INNER JOIN PPESSOA ON PFUNC.CODPESSOA = PPESSOA.CODIGO"
       
           ssql = ssql & _
                  " WHERE PFUNC.CODCOLIGADA = 1" & _
                  " AND PFUNC.DATADEMISSAO >= " & cFields(1) & ""
                  
    Case 10 'ATUALIZAR A TABELA DE PESSOAS O CAMPO PESSIST_NUMERO COM O sStatusMsg VINDO DA ACAO
           ssql = "UPDATE PESSOAS SET " & _
                  "PESSIT_NUMERO = " & cFields(18)
                  
           ssql = ssql & _
                  " WHERE PES_NUMERO = " & cFields(1)
                  
    Case 11 'ATUALIZAR A TABELA DE CREDENCIAIS O CAMPO CRED_BLOQUEADA COM O sStatusMsg VINDO DA ACAO
           ssql = "UPDATE CREDENCIAIS SET " & _
                  "CRED_BLOQUEADA = " & cFields(6)
       
           ssql = ssql & _
                  " WHERE CRED_NUMERO = " & cFields(1)

    Case 12 'Consultar CADASTRO DE PESSOAL VERIFICANDO A DATA DE PROGRAMACAO DE FERIAS
           ssql = "SELECT " & _
                  "CHAPA, " & _
                  "PFUNC.INICPROGFERIAS1, " & _
                  "PFUNC.FIMPROGFERIAS1 "
           
           ssql = ssql & _
                  "FROM PFUNC " & _
                  "INNER JOIN PPESSOA ON PFUNC.CODPESSOA = PPESSOA.CODIGO"
       
           ssql = ssql & _
                  " WHERE PFUNC.CODCOLIGADA = 1" & _
                  " AND   PFUNC.INICPROGFERIAS1 IS NOT NULL"

    Case 13 'ATUALIZAR A TABELA DE PESSOAS O CAMPO PESSIST_NUMERO,INICPROGFERIAS1,FIMPROGFERIAS1
            'COM OS DADOS VINDO DA ACAO
           ssql = "UPDATE PESSOAS SET " & _
                  "PESSIT_NUMERO = " & cFields(18) & _
                  ",PES_DTSITINICIAL = " & cFields(19) & _
                  ",PES_DTSITFINAL = " & cFields(20)
                  
           ssql = ssql & _
                  " WHERE PES_NUMERO = " & cFields(1)
                  
    Case 14 'Consultar TABELA DE MUDANCA DE SECOES EM UMA DATA DEFINIDA NO BANCO DA RM
           ssql = "SELECT " & _
                  " PFHSTSEC.CHAPA, " & _
                  " PFHSTSEC.MOTIVO, " & _
                  " CODSECAO "
           
           ssql = ssql & _
                  "FROM  PFHSTSEC "
       
           ssql = ssql & _
                  " WHERE PFHSTSEC.CODCOLIGADA = 1" & _
                  " AND PFHSTSEC.DTMUDANCA >= " & cFields(1) & ""

    Case 15 'Consultar TABELA DE ESTRUTURA DA RODBEL VERIFICANDO SE HA O FUNCIONARIO E COM A SECAO DIFERENTE. aTUALIZANDO DEPOIS
           ssql = "SELECT " & _
                  "COUNT(*) AS ENCONTRADO "
                  
'           ssql = ssql & _
'                  "(SELECT COUNT(*) FROM ESTRUTURA WHERE EST_NUMERO = " & cFields(2) & ") AS ENCONTRA_EST "
           
           ssql = ssql & _
                  "FROM PESSOAS "
       
           ssql = ssql & _
                  " WHERE PES_NUMERO = " & cFields(1) & _
                  " AND EST_NUMERO <> " & cFields(2)
    
    Case 16 'ATUALIZAR A TABELA DE PESSOAS O CAMPO EST_NUMERO, COM A NOVA SECAO DO FUNCIONARIO
           ssql = "UPDATE PESSOAS SET " & _
                  "EST_NUMERO = " & cFields(7)
                  
           ssql = ssql & _
                  " WHERE PES_NUMERO = " & cFields(1)


    Case 17 'SELECIONAR FUNCIONARIOS COM AFASTAMENTO NO PERIODO ULTIMO DIA REGISTRADO
    
           ssql = "SELECT " & _
                  "PFUNC.CHAPA, "
           ssql = ssql & _
                  "(SELECT MAX(AFAST.DTINICIO)" & _
                  " FROM PFHSTAFT AFAST " & _
                  " WHERE AFAST.CODCOLIGADA = 1 " & _
                  " AND   AFAST.CHAPA = PFUNC.CHAPA)" & _
                  " AS DTINICIO, "
          ssql = ssql & _
                 "(SELECT MAX(AFAST2.DTFINAL)" & _
                  " FROM   PFHSTAFT AFAST2 " & _
                  " WHERE  AFAST2.CODCOLIGADA = 1 " & _
                  " AND    AFAST2.CHAPA = PFUNC.CHAPA) " & _
                  " AS DTFINAL "
                  
          ssql = ssql & _
                  "FROM PFHSTAFT " & _
                  " INNER JOIN PFUNC " & _
                  " ON  PFHSTAFT.CHAPA = PFUNC.CHAPA " & _
                  " AND PFHSTAFT.CODCOLIGADA = PFUNC.CODCOLIGADA "
                  
         ssql = ssql & _
                  "WHERE (PFUNC.CODSITUACAO <> 'D') AND (PFUNC.CODCOLIGADA = 1) " & _
                  " GROUP BY PFUNC.CHAPA " & _
                  " ORDER BY PFUNC.CHAPA "
                  
    Case 18 'SELECIONAR OS GRUPOS DE ACESSO CADASTRADO NATABELA DE GRUPOS_AREAS
    
           ssql = "SELECT " & _
                  "GRP_NUMERO, " & _
                  "ARE_NUMEROENTRADA, " & _
                  "ARE_NUMEROSAIDA, " & _
                  "GRAR_TOLERANCIADESTINO, " & _
                  "GRAR_TOLERANCIAORIGEM, " & _
                  "GRAR_TIPOTOLERANCIA "
           ssql = ssql & _
                  "From GRUPOS_AREAS "
                  
    Case 19 'INCLUIR DADOS NA TABELA DE CRED_ACESSO
           
           ssql = "INSERT INTO CRED_ACESSO (" & _
                  "CRED_NUMERO, " & _
                  "GRP_NUMERO, " & _
                  "ARE_NUMERO, " & _
                  "CRACTI_NUMERO, " & _
                  "CRAC_DIFPORGRUPO, " & _
                  "JORN_NUMERO, " & _
                  "TURN_NUMERO, " & _
                  "CRAC_QTDTOTALACESSOS, " & _
                  "CRAC_QTDACESSOS, " & _
                  "CRAC_VERFSENHA, " & _
                  "CRAC_VERFDIGITAL, " & _
                  "CRAC_ULTPASSAGEM"
           
           ssql = ssql & ") VALUES (" & _
                           cFields(1) & ", " & _
                           cFields(2) & ", " & _
                           cFields(3) & ", " & _
                           cFields(4) & ", " & _
                           cFields(5) & ", " & _
                           cFields(7) & ", " & _
                           cFields(6) & ", " & _
                           cFields(8) & ", " & _
                           cFields(9) & ", " & _
                           cFields(10) & ", " & _
                           cFields(11) & ", " & _
                           cFields(12) & " )"
                  
    Case 20 'SELECIONAR FUNCIONARIOS COM ATESTADO NO PERIODO ULTIMO DIA REGISTRADO
    
           ssql = "SELECT " & _
                  "PFUNC.CHAPA,PFUNC.CODPESSOA, "
           ssql = ssql & _
                  "(SELECT MAX(ATEST.DTINICIO) " & _
                  " FROM VATESTADO ATEST " & _
                  " WHERE ATEST.CODCOLIGADA = 1 " & _
                  " AND   ATEST.CODPESSOA = PFUNC.CODPESSOA)" & _
                  " AS DTINICIO, "
          ssql = ssql & _
                 "(SELECT MAX(ATEST2.DTFINAL)" & _
                  " FROM   VATESTADO ATEST2 " & _
                  " WHERE  ATEST2.CODCOLIGADA = 1 " & _
                  " AND    ATEST2.CODPESSOA = PFUNC.CODPESSOA) " & _
                  " AS DTFINAL "
                  
          ssql = ssql & _
                  "FROM VATESTADO " & _
                  " INNER JOIN PFUNC " & _
                  " ON  VATESTADO.CODPESSOA = PFUNC.CODPESSOA " & _
                  " AND VATESTADO.CODCOLIGADA = PFUNC.CODCOLIGADA "
                  
         ssql = ssql & _
                  "WHERE (PFUNC.CODSITUACAO <> 'D') AND (PFUNC.CODCOLIGADA = 1) " & _
                  " GROUP BY PFUNC.CHAPA,PFUNC.CODPESSOA " & _
                  " ORDER BY PFUNC.CHAPA,PFUNC.CODPESSOA "
                  
    Case 21 'SELECIONAR OS SETORES DA RM
    
           ssql = "SELECT " & _
                  "PSECAO.CODCOLIGADA, " & _
                  "PSECAO.CODIGO,  " & _
                  "PSECAO.DESCRICAO "
           ssql = ssql & _
                  "FROM PSECAO "
           

    Case 22 'INCLUIR OS SETORES DA RM NA RODBEL
    
           ssql = "INSERT INTO ESTRUTURA (" & _
                  "EST_NUMERO, " & _
                  "EST_DESCRICAO, " & _
                  "EST_LOGO, " & _
                  "EST_EMPRESA"
           
           ssql = ssql & ") VALUES (" & _
                           cFields(1) & ", " & _
                           cFields(2) & ", " & _
                           cFields(3) & ", " & _
                           cFields(4) & " )"

    Case 23 'VERIFICAR EXISTENCIA DO CODIGO DO SETOR NA RODBEL
    
           ssql = "SELECT " & _
                  "ESTRUTURA.NUMERO "
           ssql = ssql & _
                  "FROM ESTRUTURA "
           ssql = ssql & _
                     "WHERE ESTRUTURA.EST_EMPRESA = " & cFields(1) & _
                     "  AND ESTRUTURA.EST_NUMERO = " & cFields(2)

End Select

Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch

rs.Open ssql, cCONECCAO

Set RegBanco_dados = rs
Set rs = Nothing

Exit Function

Erro:
Set rs = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Rem *****************************************************************************
Rem *****************************************************************************
Rem FUNCAO PARA REGISTRAR EM ARQUIVO DE LOG DO TIPO TXT ONDE O EXECUTAVEL ESTIVER
Rem AS ATUALIZACOES E ERROS QUE HOUVER NA ATUALIZACAO DOS REGISTROS
Rem *****************************************************************************
Rem *****************************************************************************
Public Function Gerar_Situacao_Log(c_fields As Collection)

Dim Y As Double
Dim nx As Integer

Open App.Path & "\LogRmRodBel.TXT" For Random Access Read Write As #11 Len = Len(sTexto)

Y = LOF(11) / Len(sTexto)

For nx = 1 To c_fields.Count

    sTexto.Texto = c_fields(nx)

    sTexto.FFinal = Chr$(13) + Chr$(10)
    Y = Y + 1
    Put 11, Y, sTexto

Next

Close #11

End Function


