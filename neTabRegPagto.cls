VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neTabRegPagto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS DO Tabela RegPagto vale
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function TabRegPagto_Consultar_Funcionario(ByVal sBanco As String, _
                                                  ByVal c_COLIGADA As String, _
                                                  ByVal c_CHAPA As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_COLIGADA, "c_COLIGADA"
cFields.Add c_CHAPA, "c_CHAPA"

Set TabRegPagto_Consultar_Funcionario = TabRegPagto_Critica(ADOConnection, cFields, 10)

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
Public Function TabRegPagto_Consultar_Funcionario_Geral(ByVal sBanco As String, _
                                                        ByVal c_COLIGADA) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_COLIGADA, "c_COLIGADA"

Set TabRegPagto_Consultar_Funcionario_Geral = TabRegPagto_Critica(ADOConnection, cFields, 101)

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

Public Function TabRegPagto_Consultar_Filial(ByVal sBanco As String, _
                                        Optional ByVal c_FILIAL As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB("1", "A")

If Len(Trim(c_FILIAL)) > 0 Then
   cFields.Add c_FILIAL, "c_FILIAL"
End If

Set TabRegPagto_Consultar_Filial = TabRegPagto_Critica(ADOConnection, cFields, 11)

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

Public Function TabRegPagto_Consul_UltMov(ByVal sBanco As String, _
                                 Optional ByVal c_TpoFecAbe As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

On Error GoTo Erro

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

If Len(Trim(c_TpoFecAbe)) > 0 Then
   cFields.Add c_TpoFecAbe, "c_TpoFecAbe"
End If

Set TabRegPagto_Consul_UltMov = TabRegPagto_Critica(ADOConnection, cFields, 1)

If TabRegPagto_Consul_UltMov.RecordCount = 0 Then
   Err.Raise 50000, Err.Source, "Registro de fechamento não encontrado (TabRegPagto),Procure o responsável!"
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

Public Function TabRegPagto_Consulta_MovFechados(ByVal sBanco As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

On Error GoTo Erro

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set TabRegPagto_Consulta_MovFechados = TabRegPagto_Critica(ADOConnection, cFields, 12)

If TabRegPagto_Consulta_MovFechados.RecordCount = 0 Then
   Err.Raise 50000, Err.Source, "Registro de fechamento não encontrado (TabRegPagto),Procure o responsável!"
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
Public Function TabRegPagto_Consulta_CentroCusto(ByVal sBanco As String, c_Classifica As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

On Error GoTo Erro

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_Classifica, "c_Classifica"

Set TabRegPagto_Consulta_CentroCusto = TabRegPagto_Critica(ADOConnection, cFields, 13)

If TabRegPagto_Consulta_CentroCusto.RecordCount = 0 Then
   Err.Raise 50000, Err.Source, "Registro de Centro de custo não encontrados,Procure o responsável!"
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
Public Function TabRegPagto_Consulta_Func_DifA(ByVal sBanco As String, _
                                                  ByVal sClassificacao As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

On Error GoTo Erro

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sClassificacao, "sClassificacao"

Set TabRegPagto_Consulta_Func_DifA = TabRegPagto_Critica(ADOConnection, cFields, 14)

If TabRegPagto_Consulta_Func_DifA.RecordCount = 0 Then
   Err.Raise 50000, Err.Source, "Registro de Funcionários não encontrados,Procure o responsável!"
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

Public Function TabRegPagto_Critica(ByVal cCONECCAO As Connection, _
                                ByVal cFields As Collection, _
                                ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction > 1 And nAction < 11 Then
   'Critica dos Campos dos Vendedor
    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
       Err.Raise 50000, Err.Source, "Código do VENDEDOR Inválido"
    End If
    If Len(Trim(cFields(2))) = 0 Then
       Err.Raise 50000, Err.Source, "Nome do VENDEDOR Inválido"
    End If
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
                Case 2, 3, 4, 5, 6, 7, 8, 9
                 If Len(Trim(cFields(nx))) = 0 Then
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
                Case 2, 3, 4, 5, 6, 7, 8, 9, 11
                    sSQL = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                
                ' Altera campos tipo data
                Case 0
                    sSQL = "#" & Format(cFields(nx), "mm/dd/yyyy") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                 'Altera campos tipo data da alteração
                Case 12
                    sSQL = "#" & Format(cFields(nx), "mm/dd/yyyy HH:MM:SS") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
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
                Case 9
'                    If InStr(cFields(nx), ",") > 0 Then
                        sSQL = Replace(cFields(nx), ".", "")
                        sSQL = Replace(sSQL, ",", ".")
                        cFields.Remove nx
                        cFields.Add sSQL, , , nx - 1
'                    End If
            End Select
        End If
    
    Next nx
End If

Set TabRegPagto_Critica = TabRegPagto_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE DADOS DO VENDEDOR
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function TabRegPagto_dados(ByVal cCONECCAO As Connection, _
                                  ByVal cFields As Collection, _
                                  ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String
Dim nseq As String



Select Case nAction

    Case 1 'Consultar TABELA DOS VALES PAGOS PARA PEGAR O ULTIMO REGISTRO
    sSQL = "SELECT TOP 1 " & _
           "TabRegPagto.sequencial, " & _
           "TabRegPagto.periodo1, " & _
           "TabRegPagto.periodo2, " & _
           "TabRegPagto.valpagto, " & _
           "TabRegPagto.datapagto, " & _
           "TabRegPagto.coligada, " & _
           "TabRegPagto.login " & _
           "FROM TabRegPagto "
           
    If cFields.Count > 0 Then
       If Len(cFields(1)) > 0 Then 'caso seja > 0 entao,pegara o registros DESTE SEQUENCIAL"
          sSQL = sSQL & "WHERE SEQUENCIA = " & cFields(1)
       End If
       sSQL = sSQL & " ORDER BY sequencial DESC"
    Else
       sSQL = sSQL & " ORDER BY sequencial DESC"
    End If

    Case 10 ' leitura no banco da rm para consulta de 1 funcionario

    sSQL = "SELECT " & _
           "DISTINCT(CHAPA), " & _
           "nome," & _
           "descricao, " & _
           "CODSITUACAO " & _
           "FROM PFUNC " & _
           "INNER JOIN PCODINTGERFUNC " & _
           "on pcodintgerfunc.codinterno = pfunc.INTEGRGERENCIAL " & _
           "and pcodintgerfunc.codcoligada = pfunc.CODCOLIGADA "
     
     sSQL = sSQL & _
                "WHERE PFUNC.CODCOLIGADA = " & cFields(1) & _
                "  AND CHAPA = " & cFields(2) & ""
           
    Case 101 ' leitura no banco da rm para consulta de todos funcionario

    sSQL = "SELECT " & _
           "DISTINCT(CHAPA), " & _
           "nome," & _
           "descricao, " & _
           "CODSITUACAO " & _
           "FROM PFUNC " & _
           "INNER JOIN PCODINTGERFUNC " & _
           "on pcodintgerfunc.codinterno = pfunc.INTEGRGERENCIAL " & _
           "and pcodintgerfunc.codcoligada = pfunc.CODCOLIGADA "
     
     sSQL = sSQL & _
                "WHERE PFUNC.CODCOLIGADA = " & cFields(1) & _
                "  AND (CODSITUACAO <> 'A' AND CODSITUACAO <> 'F' AND CODSITUACAO <> 'D') " & _
                " ORDER BY CHAPA "
    Case 11 ' leitura no bancoda rm pra receber dados bancarios e secao

    sSQL = "SELECT " & _
           "codagenciapagto, " & _
           "contapagamento, " & _
           "nome," & _
           "INTEGRGERENCIAL, " & _
           "descricao " & _
           "FROM PFUNC " & _
           "INNER JOIN PCODINTGERFUNC " & _
           "on pcodintgerfunc.codinterno = pfunc.INTEGRGERENCIAL " & _
           "WHERE  CHAPA = " & cFields(1) & " AND pcodintgerfunc.CODCOLIGADA=1"
    
    Case 12 'Consultar TABELA DE FECHAMENTO DOS VALES PAGOS PARA PEGAR TODOS OS REGISTRO
    
    sSQL = "SELECT " & _
           "TabRegPagto.sequencial, " & _
           "TabRegPagto.periodo1, " & _
           "TabRegPagto.periodo2, " & _
           "TabRegPagto.valpagto, " & _
           "TabRegPagto.datapagto, " & _
           "TabRegPagto.coligada, " & _
           "TabRegPagto.login " & _
           "FROM TabRegPagto "
           
    sSQL = sSQL & " ORDER BY sequencial DESC"
    
    Case 13 'Consultar TABELA DE CENTRO DE CUSTO DA RM
    
    sSQL = "SELECT " & _
           "DISTINCT(PCODINTGERFUNC.codinterno), " & _
           "PCODINTGERFUNC.descricao FROM PCODINTGERFUNC"
           
    If cFields(1) = "0" Then
       sSQL = sSQL & " ORDER BY codinterno"
    Else
       sSQL = sSQL & " ORDER BY descricao"
    End If
    
    Case 14 ' leitura no banco da rm para consulta de funcionarios da tela depesquisa

    sSQL = "SELECT " & _
           "DISTINCT(CHAPA), " & _
           "nome," & _
           "descricao " & _
           "FROM PFUNC " & _
           "INNER JOIN PCODINTGERFUNC " & _
           "on pcodintgerfunc.codinterno = pfunc.INTEGRGERENCIAL " & _
           "and pcodintgerfunc.codcoligada = pfunc.CODCOLIGADA " & _
           "WHERE PFUNC.CODCOLIGADA=1 AND CODSITUACAO<>'A' AND CODSITUACAO<>'F'"
           
     If cFields(1) = "0" Then
        sSQL = sSQL & " ORDER BY nome"
     ElseIf cFields(1) = "1" Then
        sSQL = sSQL & " ORDER BY descricao"
     Else
        sSQL = sSQL & " ORDER BY CHAPA"
     End If

End Select

Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch
rs.Open sSQL, cCONECCAO

Set TabRegPagto_dados = rs
Set rs = Nothing
Set rsSeq = Nothing

Exit Function

Erro:

Set rs = Nothing
Set rsSeq = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function






