VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neUniMvFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'**************** Ações Mov. de Funcionarios
'Consultar Mov. de Funcionarios = 1
'Incluir Mov. de Funcionarios = 2
'Alterar Mov. de Funcionarios = 3
'Excluir Mov. de Funcionarios = 4
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS da UNIMED REFERENTE A MOV DE FUNCIONARIOS
' ***********************************************************
' ***********************************************************
' ***********************************************************

Public Function MovFuncionario_Consultar(ByVal sBanco As String, _
                                Optional ByVal c_MFU_SEQ As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

If Len(Trim(c_MFU_SEQ)) > 0 Then
   cFields.Add c_MFU_SEQ, "c_MFU_SEQ"
End If

Set MovFuncionario_Consultar = MovFuncionario_Critica(ADOConnection, cFields, 1)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function MovFuncionario_Incluir(ByVal c_Ano_Mes As String, _
                                       ByVal c_percentual As String, _
                                       ByVal c_COLIGADA As String, _
                                       ByVal c_MFU_RS As ADODB.Recordset) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nx As Integer
Dim nSaldo As Double

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBancoUnimed, "A")


If c_MFU_RS.RecordCount > 0 Then
   
   c_MFU_RS.MoveFirst
   
   While Not c_MFU_RS.EOF
       nx = nx + 1
       
       Set cFields = New Collection
       cFields.Add "", "c_MFU_SEQ"
       cFields.Add c_COLIGADA, "c_MFU_TCO_CODIGO"
       cFields.Add c_Ano_Mes, "c_MFU_ANO_MES"
       cFields.Add c_MFU_RS!CHAPA, "c_MFU_CHAPA"
       cFields.Add Replace(c_MFU_RS!VALOR, ",", "."), "c_MFU_VALOR"
       cFields.Add c_percentual, "c_MFU_PER_DESC"
       cFields.Add Replace(c_MFU_RS!VALOR, ",", "."), "c_MFU_VALOR_DESC"
       cFields.Add Format(c_MFU_RS!DT_EVENTO, "dd/mm/yyyy"), "c_MFU_DT_MOV"
       cFields.Add sUsuario, "c_MFU_USUARIO"
       cFields.Add c_MFU_RS!TIPO, "c_MFU_TIPO"
       Call MovFuncionario_Critica(ADOConnection, cFields, 2)
       
       c_MFU_RS.MoveNext
   Wend
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function MovFuncionario_calcular(ByVal c_Ano_Mes As String, _
                                        ByVal c_COLIGADA As String, _
                                        ByVal c_MFU_RS As ADODB.Recordset) As ADODB.Recordset
 
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBancoUnimed, "A")

If c_MFU_RS.RecordCount > 0 Then
   
   c_MFU_RS.MoveFirst
   While Not c_MFU_RS.EOF
       Set cFields = New Collection
       cFields.Add c_COLIGADA, "c_MFU_TCO_CODIGO"
       cFields.Add Trim(c_MFU_RS!CHAPA), "c_MFU_CHAPA"
       cFields.Add c_Ano_Mes, "c_Ano_Mes"
       cFields.Add Replace(c_MFU_RS!DESCONTO, ",", "."), "DESCONTO"
       cFields.Add Replace(c_MFU_RS!SALDO, ",", "."), "SALDO"
       cFields.Add sUsuario, "c_MFU_USUARIO"
       cFields.Add c_MFU_RS!SITUACAO, "SITUACAO"
       Rem aqui marcos ver se deve-se tirar a geracao do registro qdo o desconto for = 0
       Rem atualizar o calculo do desconto
       Call MovFuncionario_Critica(ADOConnection, cFields, 3)
       Rem atualizar o saldo do funcionario
       Call MovFuncionario_Critica(ADOConnection, cFields, 4)
       
       c_MFU_RS.MoveNext
   Wend
End If

Rem atualizar o FECHAMENTO DO MOV. DO MES
Call MovFuncionario_Critica(ADOConnection, cFields, 7)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function MovFuncionario_Cancelar(ByVal c_Ano_Mes As String, _
                                        ByVal c_percentual As String, _
                                        ByVal c_COLIGADA As String, _
                                        ByVal c_MFU_RS As ADODB.Recordset) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nx As Integer
Dim nSaldo As Double

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBancoUnimed, "A")


If c_MFU_RS.RecordCount > 0 Then
   
   c_MFU_RS.MoveFirst
   
   While Not c_MFU_RS.EOF
       nx = nx + 1
       
       Set cFields = New Collection
       cFields.Add c_COLIGADA, "c_MFU_TCO_CODIGO"
       cFields.Add c_Ano_Mes, "c_MFU_ANO_MES"
       cFields.Add c_MFU_RS!CHAPA, "c_MFU_CHAPA"
       Call MovFuncionario_Critica(ADOConnection, cFields, 11)

       Rem VERIFICAR SE O MOV. ESTÁ FECHADO, CASO ESTEJA, IRÁ DEDUZIR O SALDO, SENÃO APENAS CANCELARÁ A DIGITACAO.
       Set MovFuncionario_Cancelar = CCTempneUniColigada.Coligada_Consultar(sBancoUnimed, c_COLIGADA)
       
       If MovFuncionario_Cancelar.RecordCount > 0 Then
          If MovFuncionario_Cancelar!TCO_MOV_ABERTO = "1" Then
             Rem atualizar o saldo do funcionario
             Set cFields = New Collection
             cFields.Add c_COLIGADA, "c_MFU_TCO_CODIGO"
             cFields.Add Trim(c_MFU_RS!CHAPA), "c_MFU_CHAPA"
             cFields.Add c_Ano_Mes, "c_Ano_Mes"
             nSaldo = c_MFU_RS!SALDO_ANT
             cFields.Add Replace(c_MFU_RS!VALOR, ",", "."), "DESCONTO"
             cFields.Add Replace(nSaldo, ",", "."), "SALDO"
             cFields.Add sUsuario, "c_MFU_USUARIO"
             cFields.Add "", "SITUACAO"
             Call MovFuncionario_Critica(ADOConnection, cFields, 4)
          End If
       End If
       c_MFU_RS.MoveNext
   Wend
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function MovFuncionario_ExisteMov(ByVal c_COLIGADA As String, _
                                         ByVal c_CHAPA As String, _
                                         ByVal c_ANOMES As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBancoUnimed, "A")

cFields.Add c_COLIGADA, "c_COLIGADA"
cFields.Add c_CHAPA, "c_CHAPA"
cFields.Add c_ANOMES, "c_ANOMES"

Set MovFuncionario_ExisteMov = MovFuncionario_Critica(ADOConnection, cFields, 5)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function MovFuncionario_CalculoMensal(ByVal c_COLIGADA As String, _
                                             ByVal c_ANOMES As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBancoUnimed, "A")

cFields.Add c_COLIGADA, "c_COLIGADA"
cFields.Add c_ANOMES, "c_ANOMES"

Set MovFuncionario_CalculoMensal = MovFuncionario_Critica(ADOConnection, cFields, 61)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function MovFuncionario_ConsMovExpMensal(ByVal c_COLIGADA As String, _
                                                ByVal c_ANOMES As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBancoUnimed, "A")

cFields.Add c_COLIGADA, "c_COLIGADA"
cFields.Add c_ANOMES, "c_ANOMES"

Set MovFuncionario_ConsMovExpMensal = MovFuncionario_Critica(ADOConnection, cFields, 8)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function MovFuncionario_ConsMesFechado(ByVal c_COLIGADA As String) As ADODB.Recordset

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

On Error GoTo Erro

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBancoUnimed, "A")

cFields.Add c_COLIGADA, "c_COLIGADA"

Set MovFuncionario_ConsMesFechado = MovFuncionario_Critica(ADOConnection, cFields, 9)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function MovFuncionario_Consulta_Saldo(ByVal c_COLIGADA As String, _
                                              ByVal c_CHAPA As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBancoUnimed, "A")

cFields.Add c_COLIGADA, "c_COLIGADA"
cFields.Add c_CHAPA, "c_CHAPA"

Set MovFuncionario_Consulta_Saldo = MovFuncionario_Critica(ADOConnection, cFields, 10)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function RMFuncionario_Consulta(ByVal c_COLIGADA As String, _
                              Optional ByVal c_CHAPA As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBancoRM, "A")

cFields.Add c_COLIGADA, "c_COLIGADA"

If Len(Trim(c_CHAPA)) > 0 Then
   cFields.Add c_CHAPA, "c_CHAPA"
End If

Set RMFuncionario_Consulta = MovFuncionario_Critica(ADOConnection, cFields, 31)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function MovFuncionario_ConsMovDigitado(ByVal c_COLIGADA As String, _
                                               ByVal c_MesAno As String) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBancoUnimed, "A")

cFields.Add c_COLIGADA, "c_COLIGADA"
cFields.Add c_MesAno, "c_MesAno"

Set MovFuncionario_ConsMovDigitado = MovFuncionario_Critica(ADOConnection, cFields, 32)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function MovFuncionario_ConsultaExtrato(ByVal c_COLIGADA As String, _
                              Optional ByVal c_CHAPA As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBancoUnimed, "A")

cFields.Add c_COLIGADA, "c_COLIGADA"

If Len(Trim(c_CHAPA)) > 0 Then
   cFields.Add c_CHAPA, "c_CHAPA"
End If

Set MovFuncionario_ConsultaExtrato = MovFuncionario_Critica(ADOConnection, cFields, 33)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function MovFuncionario_Copia_fechamento() As ADODB.Recordset

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBancoUnimed, "A")
 
Call MovFuncionario_Critica(ADOConnection, cFields, 34)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function MovFuncionario_Restaura_Copia() As ADODB.Recordset

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBancoUnimed, "A")
 
Call MovFuncionario_Critica(ADOConnection, cFields, 35)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function MovFuncionario_Critica(ByVal cCONECCAO As Connection, _
                                       ByVal cFields As Collection, _
                                       ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro

Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction = 2 Then
   'Critica dos Campos dos Mov. de Funcionarios
    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
       Err.Raise 50000, Err.Source, "Código do funcionário Inválido"
    End If
    If Len(Trim(cFields(2))) = 0 Then
       Err.Raise 50000, Err.Source, "Descricao Inválida"
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
            Case 3, 4
                 If Len(Trim(cFields(nx))) = 0 Then
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
                    sSQL = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                
                ' Altera campos tipo data
                Case 8
                    sSQL = "CONVERT(DATETIME, CONVERT(VARCHAR, '" & Format(cFields(nx), "yyyymmdd") & "', 103), 103)"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                ' Altera campos tipo data da alteração
'                Case 8
'                    Sql = "CONVERT(DATETIME, CONVERT(VARCHAR, '" & Format(cFields(nx), "yyyymmdd HH:MM:SS") & "', 108), 108)"
'                    'sName = cFields(nx).Name
'                    cFields.Remove nx
'                    cFields.Add sSql, , , nx - 1
                'Altera campos Boolean
                Case 0
                If Len(Trim(cFields(nx))) = 0 Then
                   If cFields(nx) Then
                       sSQL = "1"
                   Else
                       sSQL = "0"
                   End If
                   'sName = cFields(nx).Name
                   cFields.Remove nx
                   cFields.Add sSQL, , , nx - 1
                End If
                'Altera campos numéricos fracionários
                Case 5, 6
                    If InStr(cFields(nx), ",") > 0 Then
                        sSQL = Replace(cFields(nx), ".", "")
                        sSQL = Replace(sSQL, ",", ".")
                        cFields.Remove nx
                        cFields.Add sSQL, , , nx - 1
                    End If
            End Select
        End If
    
    Next nx
End If

Set MovFuncionario_Critica = MovFuncionario_dados(cCONECCAO, cFields, nAction)

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
Public Function MovFuncionario_dados(ByVal cCONECCAO As Connection, _
                                     ByVal cFields As Collection, _
                                     ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String

Select Case nAction
'##############################################################################################
'##############################################################################################
'#################### COMECA AQUI AS ROTINAS PARA ACESSO AO BANCO DA MOV. DE IMPORTACAO
'##############################################################################################
'##############################################################################################
    Case 1 'Consultar Mov. de Funcionarios

    sSQL = "SELECT " & _
           "CMU_MOV_FUNCIONARIO.MFU_SEQ, " & _
           "CMU_MOV_FUNCIONARIO.MFU_TCO_CODIGO, " & _
           "CMU_MOV_FUNCIONARIO.MFU_ANO_MES, " & _
           "CMU_MOV_FUNCIONARIO.MFU_CHAPA, " & _
           "CMU_MOV_FUNCIONARIO.MFU_TIPO, " & _
           "CMU_MOV_FUNCIONARIO.MFU_VALOR, " & _
           "CMU_MOV_FUNCIONARIO.MFU_PER_DESC, " & _
           "CMU_MOV_FUNCIONARIO.MFU_DT_MOV, " & _
           "CMU_MOV_FUNCIONARIO.FOR_DTA, " & _
           "CMU_MOV_FUNCIONARIO.FOR_DTI, " & _
           "CMU_MOV_FUNCIONARIO.FOR_USU "
           
    sSQL = sSQL & _
           " FROM CMU_MOV_FUNCIONARIO"

            
    If cFields.Count > 0 Then
       If Len(Trim(cFields(1))) > 0 Then
          sSQL = sSQL & " WHERE MFU_SEQ = " & cFields(1) & ""
       Else
          sSQL = sSQL & " ORDER BY MFU_TIPO,MFU_TCO_CODIGO,MFU_VALOR"
       End If
    Else
       sSQL = sSQL & " ORDER BY MFU_TIPO,MFU_TCO_CODIGO,MFU_VALOR"
    End If
    
'##############################################################################################
    Case 2 'Incluir Mov. de Funcionarios
     
     sDate = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"
'                    "MFU_SEQ, "

    sSQL = "DELETE FROM CMU_MOV_FUNCIONARIO "
    sSQL = sSQL & _
           " WHERE MFU_TCO_CODIGO = " & cFields(2) & _
           " AND   MFU_CHAPA = " & cFields(4) & "" & _
           " AND   MFU_ANO_MES = " & cFields(3) & ""

    Set rs = New ADODB.Recordset
    cCONECCAO.CursorLocation = adUseClientBatch
    rs.Open sSQL, cCONECCAO
     
    sSQL = "INSERT INTO CMU_MOV_FUNCIONARIO(" & _
                    " MFU_TCO_CODIGO ," & _
                    " MFU_ANO_MES ," & _
                    " MFU_CHAPA ," & _
                    " MFU_TIPO ," & _
                    " MFU_VALOR ," & _
                    " MFU_PER_DESC, " & _
                    " MFU_VALOR_DESC, " & _
                    " MFU_DT_MOV, " & _
                    " MFU_USUARIO "
    sSQL = sSQL & ") VALUES (" & _
                    cFields(2) & ", " & _
                    cFields(3) & ", " & _
                    cFields(4) & ", " & _
                    cFields(10) & ", " & _
                    cFields(5) & "," & _
                    cFields(6) & "," & _
                    cFields(7) & "," & _
                    cFields(8) & "," & _
                    cFields(9) & ")"

'##############################################################################################
    Case 3 'Atualizar o desconto a ser realizado neste mes/ano do Mov. de Funcionarios
    
    Rem VERIFICAR SE EXISTE O REGISTRO DE MOVIMENTACAO, CASO NAO TENHA É PORQUE DEVE SER SOMETE SALDO
    sDate = "'" & Format(Now(), "yyyymmdd") & "'"
    
    sSQL = "SELECT COUNT(*) AS CONTADOR FROM CMU_MOV_FUNCIONARIO "
    
    sSQL = sSQL & _
           " WHERE MFU_TCO_CODIGO = " & cFields(1) & _
           " AND MFU_CHAPA = '" & cFields(2) & "'" & _
           " AND MFU_ANO_MES = '" & cFields(3) & "'"
    
    Set rs = New ADODB.Recordset
    cCONECCAO.CursorLocation = adUseClientBatch
    rs.Open sSQL, cCONECCAO
    
    If rs!CONTADOR = 0 Then
       sSQL = "INSERT INTO CMU_MOV_FUNCIONARIO(" & _
                       " MFU_TCO_CODIGO ," & _
                       " MFU_ANO_MES ," & _
                       " MFU_CHAPA ," & _
                       " MFU_TIPO ," & _
                       " MFU_VALOR ," & _
                       " MFU_PER_DESC, " & _
                       " MFU_VALOR_DESC, " & _
                       " MFU_DT_MOV, " & _
                       " MFU_USUARIO,MFU_DT_EXP,MFU_SITUACAO "
       sSQL = sSQL & ") VALUES ('" & _
                       cFields(1) & "','" & _
                       cFields(3) & "','" & _
                       cFields(2) & "', " & _
                       "0, " & _
                       "0, " & _
                       "0," & _
                       cFields(4) & "," & _
                       sDate & "," & _
                       cFields(6) & "," & sDate & ",'" & cFields(7) & "')"
    
    Else

       sSQL = "UPDATE CMU_MOV_FUNCIONARIO SET" & _
                    " MFU_VALOR_DESC = " & cFields(4) & _
                    ",MFU_DT_EXP = " & sDate & _
                    ",MFU_USU_EXP = " & cFields(6) & _
                    ",MFU_SITUACAO = '" & cFields(7) & "'"
                    
       sSQL = sSQL & _
              " WHERE MFU_TCO_CODIGO = " & cFields(1) & _
              " AND MFU_CHAPA = '" & cFields(2) & "'" & _
              " AND MFU_ANO_MES = '" & cFields(3) & "'"
    End If

'##############################################################################################
    Case 4 'Atualizar o SALDO neste mes/ano do Mov. de Funcionarios na tabela de saldo
    
    sDate = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"
    
    Rem verificar a existencia da tabela CMU_TAB_SALDO
    
    sSQL = "SELECT " & _
           "CMU_TAB_SALDO.SAL_SALDO " & _
           "FROM CMU_TAB_SALDO "
    sSQL = sSQL & _
           " WHERE SAL_TCO_COLIGADA = " & cFields(1) & _
           " AND SAL_CHAPA = '" & cFields(2) & "'"
           
    Set rs = New ADODB.Recordset
    cCONECCAO.CursorLocation = adUseClientBatch
    rs.Open sSQL, cCONECCAO
    
    If rs.RecordCount = 0 Then
       sSQL = "INSERT INTO CMU_TAB_SALDO(" & _
                       " SAL_TCO_COLIGADA ," & _
                       " SAL_CHAPA ," & _
                       " SAL_SALDO "
       sSQL = sSQL & ") VALUES (" & _
                       cFields(1) & ",'" & _
                       cFields(2) & "', " & _
                       cFields(5) & ")"
    Else
       sSQL = "UPDATE CMU_TAB_SALDO SET" & _
                    " SAL_SALDO = " & cFields(5)
                    
       sSQL = sSQL & _
              " WHERE SAL_TCO_COLIGADA = " & cFields(1) & _
              " AND SAL_CHAPA = '" & cFields(2) & "'"
    
    End If
    
    
'##############################################################################################
    Case 5 'Consultar Existencia do Mov. de Funcionarios, pela empresa,chapa e ano/mes do movimento

    sSQL = "SELECT " & _
           "CMU_MOV_FUNCIONARIO.MFU_SEQ, " & _
           "CMU_MOV_FUNCIONARIO.MFU_TCO_CODIGO, " & _
           "CMU_MOV_FUNCIONARIO.MFU_ANO_MES, " & _
           "CMU_MOV_FUNCIONARIO.MFU_CHAPA, " & _
           "CMU_MOV_FUNCIONARIO.MFU_TIPO, " & _
           "CMU_MOV_FUNCIONARIO.MFU_VALOR, " & _
           "CMU_MOV_FUNCIONARIO.MFU_PER_DESC, " & _
           "CMU_MOV_FUNCIONARIO.MFU_DT_MOV "
           
    sSQL = sSQL & _
           " FROM CMU_MOV_FUNCIONARIO"
    sSQL = sSQL & _
           " WHERE MFU_TCO_CODIGO = " & cFields(1) & _
           " AND MFU_CHAPA = '" & cFields(2) & "'" & _
           " AND MFU_ANO_MES = '" & cFields(3) & "'"

'##############################################################################################
    Case 6 'Consultar Existencia do Mov. de Funcionarios, pela empresa e ano/mes do movimento p/ calculo

    sSQL = "SELECT " & _
           "CMU_MOV_FUNCIONARIO.MFU_SEQ, " & _
           "CMU_MOV_FUNCIONARIO.MFU_TCO_CODIGO, " & _
           "CMU_MOV_FUNCIONARIO.MFU_ANO_MES, " & _
           "CMU_MOV_FUNCIONARIO.MFU_CHAPA, " & _
           "CMU_MOV_FUNCIONARIO.MFU_TIPO, " & _
           "CMU_TAB_SALDO.SAL_SALDO, " & _
           "CMU_MOV_FUNCIONARIO.MFU_VALOR, " & _
           "CMU_MOV_FUNCIONARIO.MFU_PER_DESC, " & _
           "CMU_MOV_FUNCIONARIO.MFU_DT_MOV "
           
    sSQL = sSQL & _
           " FROM CMU_MOV_FUNCIONARIO "
    sSQL = sSQL & _
           "LEFT OUTER JOIN CMU_TAB_SALDO " & _
           "ON   CMU_MOV_FUNCIONARIO.MFU_TCO_CODIGO = CMU_TAB_SALDO.SAL_TCO_COLIGADA " & _
           "AND  CMU_MOV_FUNCIONARIO.MFU_CHAPA = CMU_TAB_SALDO.SAL_CHAPA"
    sSQL = sSQL & _
           " WHERE MFU_TCO_CODIGO = " & cFields(1) & _
           " AND MFU_ANO_MES = '" & cFields(2) & "'"
    
'##############################################################################################
    Case 61 'Consultar Existencia do Mov. de Funcionarios, pela empresa e ano/mes do movimento p/ calculo
            'UNINDO COM O MOVIMENTO DAS PESSOAS QUE SO TEM SALDO.
            
    sSQL = "SELECT " & _
           "CMU_MOV_FUNCIONARIO.MFU_SEQ, " & _
           "CMU_MOV_FUNCIONARIO.MFU_TCO_CODIGO, " & _
           "CMU_MOV_FUNCIONARIO.MFU_ANO_MES, " & _
           "CMU_MOV_FUNCIONARIO.MFU_CHAPA, " & _
           "CMU_MOV_FUNCIONARIO.MFU_TIPO, " & _
           "CMU_TAB_SALDO.SAL_SALDO, " & _
           "CMU_MOV_FUNCIONARIO.MFU_VALOR, " & _
           "CMU_MOV_FUNCIONARIO.MFU_PER_DESC, " & _
           "CMU_MOV_FUNCIONARIO.MFU_DT_MOV "
           
    sSQL = sSQL & _
           " FROM CMU_MOV_FUNCIONARIO "
    sSQL = sSQL & _
           "LEFT OUTER JOIN CMU_TAB_SALDO " & _
           "ON   CMU_MOV_FUNCIONARIO.MFU_TCO_CODIGO = CMU_TAB_SALDO.SAL_TCO_COLIGADA " & _
           "AND  CMU_MOV_FUNCIONARIO.MFU_CHAPA = CMU_TAB_SALDO.SAL_CHAPA"
    sSQL = sSQL & _
           " WHERE MFU_TCO_CODIGO = " & cFields(1) & _
           " AND MFU_ANO_MES = '" & cFields(2) & "'"
    
'    sSQL = sSQL & " ORDER BY MFU_CHAPA "
    
    sSQL = sSQL & " UNION "
    
    sSQL = sSQL & "SELECT " & _
           "0 AS MFU_SEQ, " & _
           "SAL_TCO_COLIGADA, "
    sSQL = sSQL & _
           "(SELECT TCO_ANO_MES_PROC From CMU_COLIGADA " & _
           "WHERE  TCO_CODIGO = " & cFields(1) & " ) AS MFU_ANO_MES, "
    sSQL = sSQL & _
           "SAL_CHAPA AS MFU_CHAPA, " & _
           "0 AS MFU_TIPO, " & _
           "SAL_SALDO, " & _
           "0 AS MFU_VALOR, "
    sSQL = sSQL & _
           "(SELECT TCO_DESCONTO FROM CMU_COLIGADA " & _
           "WHERE  TCO_CODIGO = " & cFields(1) & " ) AS MFU_PER_DESC, "
    sSQL = sSQL & _
           "0 AS MFU_DT_MOV "

    sSQL = sSQL & _
           "FROM CMU_TAB_SALDO " & _
           "WHERE SAL_SALDO > 0 AND " & _
           "      SAL_CHAPA NOT IN(SELECT MFU_CHAPA " & _
                                   " FROM CMU_MOV_FUNCIONARIO " & _
                                   " WHERE MFU_ANO_MES = '" & cFields(2) & "')"
'    sSQL = sSQL & " ORDER BY SAL_CHAPA "
'##############################################################################################
    Case 7 'Alterar a coligada para fechada
    sSQL = "UPDATE CMU_COLIGADA SET" & _
                    " TCO_MOV_ABERTO = 1"
    sSQL = sSQL & _
                 "  WHERE CMU_COLIGADA.TCO_CODIGO = " & cFields(1)

'##############################################################################################
    Case 8 'Consultar Mov. de Funcionarios fechados e prontos para exportacao, pela empresa e ano/mes do movimento p/ calculado

    sSQL = "SELECT " & _
           "CMU_MOV_FUNCIONARIO.MFU_SEQ, " & _
           "CMU_MOV_FUNCIONARIO.MFU_TCO_CODIGO, " & _
           "CMU_MOV_FUNCIONARIO.MFU_ANO_MES, " & _
           "CMU_MOV_FUNCIONARIO.MFU_CHAPA, " & _
           "CMU_MOV_FUNCIONARIO.MFU_TIPO, " & _
           "CMU_TAB_SALDO.SAL_SALDO, " & _
           "CMU_MOV_FUNCIONARIO.MFU_VALOR, " & _
           "CMU_MOV_FUNCIONARIO.MFU_VALOR_DESC, " & _
           "CMU_MOV_FUNCIONARIO.MFU_DT_MOV "
           
    sSQL = sSQL & _
           " FROM CMU_MOV_FUNCIONARIO "
    sSQL = sSQL & _
           "LEFT OUTER JOIN CMU_TAB_SALDO " & _
           "ON   CMU_MOV_FUNCIONARIO.MFU_TCO_CODIGO = CMU_TAB_SALDO.SAL_TCO_COLIGADA " & _
           "AND  CMU_MOV_FUNCIONARIO.MFU_CHAPA = CMU_TAB_SALDO.SAL_CHAPA "
    sSQL = sSQL & _
           " WHERE MFU_VALOR_DESC > 0 " & _
           " AND   MFU_TCO_CODIGO = " & cFields(1) & _
           " AND   MFU_ANO_MES = '" & cFields(2) & "'"
           
    sSQL = sSQL & " ORDER BY MFU_CHAPA"

'##############################################################################################
    Case 9 'Consultar O MES/ANO DOS Mov. os fechados DE UMA DETERMINADA EMPRESA

    sSQL = "SELECT " & _
           "CMU_MOV_FUNCIONARIO.MFU_ANO_MES "
    sSQL = sSQL & _
           " FROM CMU_MOV_FUNCIONARIO "
    sSQL = sSQL & _
           " WHERE MFU_TCO_CODIGO = " & cFields(1)
    sSQL = sSQL & _
           " GROUP BY MFU_ANO_MES "
    sSQL = sSQL & _
           " ORDER BY MFU_ANO_MES DESC"

'##############################################################################################
    Case 10 'Consultar o SALDO dos Funcionarios na tabela de saldo
    
    Rem verificar a existencia da tabela CMU_TAB_SALDO
    
    sSQL = "SELECT " & _
           "CMU_TAB_SALDO.SAL_CHAPA, " & _
           "CMU_TAB_SALDO.SAL_SALDO " & _
           "FROM CMU_TAB_SALDO "
    sSQL = sSQL & _
           " WHERE SAL_TCO_COLIGADA = " & cFields(1)
    If Len(cFields(2)) > 0 Then
       sSQL = sSQL & _
              " AND SAL_CHAPA = '" & cFields(2) & "'"
    End If

'##############################################################################################
    Case 11 'Remover o Mov de digitacao avulsa na tabela de movimentacao
    
    sSQL = "DELETE FROM CMU_MOV_FUNCIONARIO "
    sSQL = sSQL & _
           " WHERE " & _
           " MFU_TCO_CODIGO = " & cFields(1) & _
           " AND   MFU_CHAPA = '" & cFields(3) & "'" & _
           " AND   MFU_ANO_MES = '" & cFields(2) & "'"
           
'##############################################################################################
'##############################################################################################
'#################### COMECA AQUI AS ROTINAS PARA ACESSO AO BANCO DA RM, DE 31 A 50
'##############################################################################################
'##############################################################################################

    Case 31 'Consultar CADASTRO DE PESSOAL VERIFICANDO A DATA DE ADMISSAO
           sSQL = "SELECT " & _
                  "CONVERT(BIGINT,PFUNC.CHAPA) AS CHAPA, " & _
                  "PFUNC.NOME, " & _
                  "PFUNC.INTEGRGERENCIAL, " & _
                  "PFUNC.CODSITUACAO, " & _
                  "PPESSOA.CPF, " & _
                  "PPESSOA.CARTIDENTIDADE, " & _
                  "PFUNC.CODFUNCAO, " & _
                  "PFUNCAO.CARGO, " & _
                  "PFUNC.SALARIO, " & _
                  "PFUNC.CODSITUACAO AS CODSITUACAO, " & _
                  "PFUNC.DATAADMISSAO, " & _
                  "PFUNC.DATADEMISSAO, " & _
                  "PFUNC.DTAPOSENTADORIA "
           
           sSQL = sSQL & _
                  "FROM PFUNC " & _
                  "INNER JOIN PPESSOA ON PFUNC.CODPESSOA = PPESSOA.CODIGO " & _
                  "INNER JOIN PFUNCAO ON PFUNC.CODFUNCAO = PFUNCAO.CODIGO AND PFUNC.CODCOLIGADA = PFUNCAO.CODCOLIGADA "
       
           sSQL = sSQL & _
                  " WHERE PFUNC.CODCOLIGADA = " & cFields(1) & _
                  " AND PFUNC.CHAPA = " & cFields(2)
           
'##############################################################################################
'#################### SELECIONAR AS DIGITACAOES FEITAS MANUALMENTE REFERENTE AO PERIODO DESEJADO
'##############################################################################################

    Case 32

           sSQL = "SELECT " & _
                  "CMU_MOV_FUNCIONARIO.MFU_SEQ, " & _
                  "CMU_MOV_FUNCIONARIO.MFU_TCO_CODIGO, " & _
                  "CMU_MOV_FUNCIONARIO.MFU_ANO_MES, " & _
                  "CMU_MOV_FUNCIONARIO.MFU_CHAPA, " & _
                  "CMU_MOV_FUNCIONARIO.MFU_TIPO, " & _
                  "CMU_TAB_SALDO.SAL_SALDO, " & _
                  "CMU_MOV_FUNCIONARIO.MFU_VALOR, " & _
                  "CMU_MOV_FUNCIONARIO.MFU_VALOR_DESC, " & _
                  "CMU_MOV_FUNCIONARIO.MFU_DT_MOV "
                  
           sSQL = sSQL & _
                  " FROM CMU_MOV_FUNCIONARIO "
           sSQL = sSQL & _
                  "LEFT OUTER JOIN CMU_TAB_SALDO " & _
                  "ON   CMU_MOV_FUNCIONARIO.MFU_TCO_CODIGO = CMU_TAB_SALDO.SAL_TCO_COLIGADA " & _
                  "AND  CMU_MOV_FUNCIONARIO.MFU_CHAPA = CMU_TAB_SALDO.SAL_CHAPA "
                  
Rem marcos 14/02/2014, antes " AND   MFU_TIPO = 1 "
           sSQL = sSQL & _
                  " WHERE MFU_ANO_MES = '" & cFields(2) & "'" & _
                  " AND   MFU_TIPO <> 0 "
           If cFields(1) <> Empty Then
              sSQL = sSQL & _
                  " AND   MFU_TCO_CODIGO = " & cFields(1)
           End If
           
           sSQL = sSQL & _
                  " ORDER BY MFU_CHAPA"

'##############################################################################################
'#################### Consultar extrato de Funcionarios
'##############################################################################################
    Case 33

    sSQL = "SELECT " & _
           "CMU_MOV_FUNCIONARIO.MFU_ANO_MES, " & _
           "CMU_MOV_FUNCIONARIO.MFU_TIPO, " & _
           "CMU_MOV_FUNCIONARIO.MFU_VALOR, " & _
           "CMU_MOV_FUNCIONARIO.MFU_PER_DESC, " & _
           "CMU_MOV_FUNCIONARIO.MFU_DT_MOV "
           
    sSQL = sSQL & _
           " FROM CMU_MOV_FUNCIONARIO"

    sSQL = sSQL & " WHERE MFU_CHAPA = '" & cFields(2) & "'"
    sSQL = sSQL & " ORDER BY MFU_ANO_MES DESC"
                  
       
'##############################################################################################
'#################### SALVAR A TABELA DE MOVIMENTO E DE SALDO EM OUTRA AUXILIAR ANTES DE FECHAR O MES
'##############################################################################################
    Case 34
    
        sSQL = "DELETE FROM CMU_TAB_SALDO_ANT"
                  
        Set rs = New ADODB.Recordset
        cCONECCAO.CursorLocation = adUseClientBatch
        rs.Open sSQL, cCONECCAO
        
        sSQL = "INSERT INTO CMU_TAB_SALDO_ANT(" & _
                        " SAL_TCO_COLIGADA ," & _
                        " SAL_CHAPA ," & _
                        " SAL_SALDO) "
        sSQL = sSQL & "SELECT " & _
                        " SAL_TCO_COLIGADA ," & _
                        " SAL_CHAPA ," & _
                        " SAL_SALDO FROM  CMU_TAB_SALDO "
                  
        Set rs = New ADODB.Recordset
        cCONECCAO.CursorLocation = adUseClientBatch
        rs.Open sSQL, cCONECCAO

        sSQL = "DELETE FROM CMU_MOV_FUNCIONARIO_ANT"
                  
        Set rs = New ADODB.Recordset
        cCONECCAO.CursorLocation = adUseClientBatch
        rs.Open sSQL, cCONECCAO
        
        sSQL = "INSERT INTO CMU_MOV_FUNCIONARIO_ANT(" & _
                        " MFU_TCO_CODIGO ," & _
                        " MFU_ANO_MES ," & _
                        " MFU_CHAPA ," & _
                        " MFU_TIPO ," & _
                        " MFU_VALOR ," & _
                        " MFU_PER_DESC, " & _
                        " MFU_VALOR_DESC, " & _
                        " MFU_DT_MOV, " & _
                        " MFU_USUARIO,MFU_DT_EXP,MFU_SITUACAO) "

        sSQL = sSQL & "SELECT" & _
                        " MFU_TCO_CODIGO ," & _
                        " MFU_ANO_MES ," & _
                        " MFU_CHAPA ," & _
                        " MFU_TIPO ," & _
                        " MFU_VALOR ," & _
                        " MFU_PER_DESC, " & _
                        " MFU_VALOR_DESC, " & _
                        " MFU_DT_MOV, " & _
                        " MFU_USUARIO,MFU_DT_EXP,MFU_SITUACAO FROM  CMU_MOV_FUNCIONARIO "

'##############################################################################################
'#################### RESTAURAR TABELA DE MOVIMENTO E DE SALDO EM OUTRA AUXILIAR
'##############################################################################################
    Case 35
    
        sSQL = "DELETE FROM CMU_TAB_SALDO"
                  
        Set rs = New ADODB.Recordset
        cCONECCAO.CursorLocation = adUseClientBatch
        rs.Open sSQL, cCONECCAO
        
        sSQL = "INSERT INTO CMU_TAB_SALDO(" & _
                        " SAL_TCO_COLIGADA ," & _
                        " SAL_CHAPA ," & _
                        " SAL_SALDO) "
        sSQL = sSQL & "SELECT " & _
                        " SAL_TCO_COLIGADA ," & _
                        " SAL_CHAPA ," & _
                        " SAL_SALDO FROM  CMU_TAB_SALDO_ANT "
                  
        Set rs = New ADODB.Recordset
        cCONECCAO.CursorLocation = adUseClientBatch
        rs.Open sSQL, cCONECCAO

        sSQL = "DELETE FROM CMU_MOV_FUNCIONARIO"
                  
        Set rs = New ADODB.Recordset
        cCONECCAO.CursorLocation = adUseClientBatch
        rs.Open sSQL, cCONECCAO
        
        sSQL = "INSERT INTO CMU_MOV_FUNCIONARIO(" & _
                        " MFU_TCO_CODIGO ," & _
                        " MFU_ANO_MES ," & _
                        " MFU_CHAPA ," & _
                        " MFU_TIPO ," & _
                        " MFU_VALOR ," & _
                        " MFU_PER_DESC, " & _
                        " MFU_VALOR_DESC, " & _
                        " MFU_DT_MOV, " & _
                        " MFU_USUARIO,MFU_DT_EXP,MFU_SITUACAO) "

        sSQL = sSQL & "SELECT" & _
                        " MFU_TCO_CODIGO ," & _
                        " MFU_ANO_MES ," & _
                        " MFU_CHAPA ," & _
                        " MFU_TIPO ," & _
                        " MFU_VALOR ," & _
                        " MFU_PER_DESC, " & _
                        " MFU_VALOR_DESC, " & _
                        " MFU_DT_MOV, " & _
                        " MFU_USUARIO,MFU_DT_EXP,MFU_SITUACAO FROM  CMU_MOV_FUNCIONARIO_ANT "

End Select

Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch

rs.Open sSQL, cCONECCAO

Set MovFuncionario_dados = rs
Set rs = Nothing
Set rsSeq = Nothing

Exit Function

Erro:
'cCONECCAO.RollbackTrans: cCONECCAO.BeginTrans
''cCONECCAO.Close
Set rs = Nothing
'Set cCONECCAO = Nothing
If Err.Number = -2147467259 Then
   Err.Description = "Erro no Banco, Sem permissão de Acesso no Momento"
End If
If Err.Number = -2147217873 Then
   Err.Description = "Erro de movimentação já cadastrada, Sem permissão de atualização no Momento"
End If
Err.Raise Err.Number, Err.Source, Err.Description
End Function


'''Public Function MovFuncionario_Alterar(ByVal sBanco As String, _
'''                                       ByVal c_MFU_SEQ As String, _
'''                                       ByVal c_MFU_TCO_CODIGO As String, _
'''                                       ByVal c_MFU_ANO_MES As String, _
'''                                       ByVal c_MFU_CHAPA As String, _
'''                                       ByVal c_MFU_TIPO As String, _
'''                                       ByVal c_MFU_VALOR As String, _
'''                                       ByVal c_MFU_DT_MOV As String, _
'''                                       ByVal c_MFU_PER_DESC As String, _
'''                                       ByVal c_MFU_USUARIO As String, _
'''                                       ByVal c_DATA_ALT As String) As ADODB.Recordset
'''
'''On Error GoTo Erro
'''Dim ADOConnection As ADODB.Connection
'''Dim cFields As Collection
'''Dim cConect As daAbertura
'''
'''
'''Set ADOConnection = New ADODB.Connection
'''Set cConect = New daAbertura
'''Set cFields = New Collection
'''Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")
'''
'''Set cFields = New Collection
'''cFields.Add c_MFU_SEQ, "c_MFU_SEQ" '1
'''cFields.Add c_MFU_TCO_CODIGO, "c_MFU_TCO_CODIGO" '2
'''cFields.Add c_MFU_ANO_MES, "c_MFU_ANO_MES" '3
'''cFields.Add c_MFU_CHAPA, "c_MFU_CHAPA" '4
'''cFields.Add c_MFU_TIPO, "c_MFU_TIPO" '5
'''cFields.Add c_MFU_VALOR, "c_MFU_VALOR" '6
'''cFields.Add c_MFU_DT_MOV, "c_MFU_DT_MOV" '7
'''cFields.Add c_MFU_PER_DESC, "c_MFU_PER_DESC" '8
'''cFields.Add c_MFU_USUARIO, "c_MFU_USUARIO" '9
'''cFields.Add c_DATA_ALT, "c_DATA_ALT" '10
'''
'''Set MovFuncionario_Alterar = MovFuncionario_Critica(ADOConnection, cFields, 3)
'''
'''Call CCTempneMFU_USUARIO.MFU_USUARIO_Alterar_Acesso(ADOConnection, cFields)
'''
'''ADOConnection.CommitTrans
'''Set ADOConnection = Nothing
'''Set cConect = Nothing
'''Set cFields = Nothing
'''
'''Exit Function
'''
'''Erro:
'''If ADOConnection <> "" Then
'''   ADOConnection.RollbackTrans
'''End If
'''Set ADOConnection = Nothing
'''Set cConect = Nothing
'''Set cFields = Nothing
'''Err.Raise Err.Number, Err.Source, Err.Description
'''End Function
'''Public Function MovFuncionario_Excluir(ByVal sBanco As String, _
'''                                       ByVal c_MFU_SEQ As String, _
'''                                       ByVal c_MFU_USUARIO As String, _
'''                                       ByVal c_DATA_ALT As String) As ADODB.Recordset
'''On Error GoTo Erro
'''Dim ADOConnection As ADODB.Connection
'''Dim cFields As Collection
'''Dim cConect As daAbertura
'''
'''
'''Set ADOConnection = New ADODB.Connection
'''Set cConect = New daAbertura
'''Set cFields = New Collection
'''Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")
'''
'''Set cFields = New Collection
'''cFields.Add c_MFU_SEQ, "c_MFU_SEQ" '1
'''cFields.Add Null, "c_MFU_TCO_CODIGO" '2
'''cFields.Add Null, "c_MFU_ANO_MES" '3
'''cFields.Add Null, "c_MFU_CHAPA" '4
'''cFields.Add Null, "c_MFU_TIPO" '5
'''cFields.Add Null, "c_MFU_VALOR" '6
'''cFields.Add Null, "c_MFU_DT_MOV" '7
'''cFields.Add Null, "c_MFU_PER_DESC" '8
'''cFields.Add c_MFU_USUARIO, "c_MFU_USUARIO" '9
'''cFields.Add c_DATA_ALT, "c_DATA_ALT" '10
'''
'''Set MovFuncionario_Excluir = MovFuncionario_Critica(ADOConnection, cFields, 4)
'''
'''ADOConnection.CommitTrans
'''Set ADOConnection = Nothing
'''Set cConect = Nothing
'''Set cFields = Nothing
'''
'''Exit Function
'''
'''Erro:
'''If ADOConnection <> "" Then
'''   ADOConnection.RollbackTrans
'''End If
'''Set ADOConnection = Nothing
'''Set cConect = Nothing
'''Set cFields = Nothing
'''Err.Raise Err.Number, Err.Source, Err.Description
'''End Function
'''



'''    Case 3 'Alterar Mov. de Funcionarios
'''
'''    sDate = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"
''''                 "  CMU_MOV_FUNCIONARIO.MFU_SEQ = " & cFields(1)
'''    sSql = "UPDATE CMU_MOV_FUNCIONARIO SET" & _
'''                    "  MFU_TCO_CODIGO = " & cFields(2) & _
'''                    " ,MFU_ANO_MES = " & cFields(3) & _
'''                    " ,MFU_CHAPA = " & cFields(4) & _
'''                    " ,MFU_TIPO = " & cFields(5) & _
'''                    " ,MFU_VALOR = " & cFields(6) & _
'''                    " ,MFU_PER_DESC = " & cFields(7) & _
'''                    " ,MFU_DT_MOV = " & cFields(8) & _
'''                    " WHERE MFU_SEQ = " & cFields(1)
'''
'''    Case 4 'Excluir Mov. de Funcionarios
'''    sSql = "DELETE  FROM CMU_MOV_FUNCIONARIO" & _
'''                 "  WHERE CMU_MOV_FUNCIONARIO.MFU_SEQ = " & cFields(1) & _
'''                 " AND CMU_MOV_FUNCIONARIO.FOR_DTA = " & cFields(8)

