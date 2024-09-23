VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neUniColigada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'**************** Ações Empresas Coligadas
'Consultar Empresas Coligadas = 1
'Incluir Empresas Coligadas = 2
'Alterar Empresas Coligadas = 3
'Excluir Empresas Coligadas = 4
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS DO Feriado
' ***********************************************************
' ***********************************************************
' ***********************************************************

Public Function Coligada_Incluir(ByVal sBanco As String, _
                                 ByVal c_TCO_CODIGO As String, _
                                 ByVal c_TCO_DESCRICAO As String, _
                                 ByVal c_TCO_DESCONTO As String, _
                                 ByVal c_TCO_DT_MOD As String, _
                                 ByVal c_TCO_VERBA As String, _
                                 ByVal c_TCO_ANO_MES_PROC As String, _
                                 ByVal c_TCO_MOV_ABERTO As String) As ADODB.Recordset

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_TCO_CODIGO, "c_TCO_CODIGO"
cFields.Add c_TCO_DESCRICAO, "c_TCO_DESCRICAO"
cFields.Add c_TCO_DESCONTO, "c_TCO_DESCONTO"
cFields.Add c_TCO_DT_MOD, "c_TCO_DT_MOD"
cFields.Add c_TCO_VERBA, "c_TCO_VERBA"
cFields.Add c_TCO_ANO_MES_PROC, "c_TCO_ANO_MES_PROC"
cFields.Add c_TCO_MOV_ABERTO, "c_TCO_MOV_ABERTO"

Set Coligada_Incluir = Coligada_Critica(ADOConnection, cFields, 2)

Set cFields = New Collection

cFields.Add Coligada_Incluir.Fields(0), "c_TCO_CODIGO"
cFields.Add c_TCO_DESCRICAO, "c_TCO_DESCRICAO"
cFields.Add c_TCO_DESCONTO, "c_TCO_DESCONTO"
cFields.Add c_TCO_DT_MOD, "c_TCO_DT_MOD"
cFields.Add c_TCO_VERBA, "c_TCO_VERBA"
cFields.Add c_TCO_ANO_MES_PROC, "c_TCO_ANO_MES_PROC"
cFields.Add c_TCO_MOV_ABERTO, "c_TCO_MOV_ABERTO"

Call CCTempneTCO_MOV_ABERTO.TCO_MOV_ABERTO_Inserir_Acesso(ADOConnection, cFields)

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

Public Function Coligada_Consultar(ByVal sBanco As String, _
                          Optional ByVal c_TCO_CODIGO As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

If Len(Trim(c_TCO_CODIGO)) > 0 Then
   cFields.Add c_TCO_CODIGO, "c_TCO_CODIGO"
End If

Set Coligada_Consultar = Coligada_Critica(ADOConnection, cFields, 1)

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
Public Function Coligada_Alterar(ByVal sBanco As String, _
                                 ByVal c_TCO_CODIGO As String, _
                                 ByVal c_TCO_DESCRICAO As String, _
                                 ByVal c_TCO_DESCONTO As String, _
                                 ByVal c_TCO_DT_MOD As String, _
                                 ByVal c_TCO_VERBA As String, _
                                 ByVal c_TCO_ANO_MES_PROC As String, _
                                 ByVal c_TCO_MOV_ABERTO)

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set cFields = New Collection
cFields.Add c_TCO_CODIGO, "c_TCO_CODIGO" '1
cFields.Add c_TCO_DESCRICAO, "c_TCO_DESCRICAO" '2
cFields.Add c_TCO_DESCONTO, "c_TCO_DESCONTO" '3
cFields.Add c_TCO_DT_MOD, "c_TCO_DT_MOD" '4
cFields.Add c_TCO_VERBA, "c_TCO_VERBA" '5
cFields.Add c_TCO_ANO_MES_PROC, "c_TCO_ANO_MES_PROC" '6
cFields.Add c_TCO_MOV_ABERTO, "c_TCO_MOV_ABERTO" '7

Call Coligada_Critica(ADOConnection, cFields, 3)

'Call CCTempneTCO_MOV_ABERTO.TCO_MOV_ABERTO_Alterar_Acesso(ADOConnection, cFields)

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
Public Function Coligada_Excluir(ByVal sBanco As String, _
                                 ByVal c_TCO_CODIGO As String, _
                                 ByVal c_TCO_MOV_ABERTO As String, _
                                 ByVal c_DATA_ALT As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set cFields = New Collection
cFields.Add c_TCO_CODIGO, "c_TCO_CODIGO" '1
cFields.Add Null, "c_TCO_DESCRICAO" '2
cFields.Add Null, "c_TCO_DESCONTO" '3
cFields.Add Null, "c_TCO_DT_MOD" '4
cFields.Add Null, "c_TCO_VERBA" '5
cFields.Add Null, "c_TCO_ANO_MES_PROC" '6
cFields.Add c_TCO_MOV_ABERTO, "c_TCO_MOV_ABERTO" '7
cFields.Add c_DATA_ALT, "c_DATA_ALT" '8

Set Coligada_Excluir = Coligada_Critica(ADOConnection, cFields, 4)

Call CCTempneTCO_MOV_ABERTO.TCO_MOV_ABERTO_Excluir_Acesso(ADOConnection, cFields)

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
Public Function Coligada_Fechamento(ByVal sBanco As String, _
                                    ByVal c_TCO_CODIGO) As ADODB.Recordset

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set cFields = New Collection
cFields.Add c_TCO_CODIGO, "c_TCO_CODIGO" '1
cFields.Add Null, "c_TCO_DESCRICAO" '2
cFields.Add Null, "c_TCO_DESCONTO" '3
cFields.Add Null, "c_TCO_DT_MOD" '4
cFields.Add Null, "c_TCO_VERBA" '5
cFields.Add Null, "c_TCO_ANO_MES_PROC" '6
cFields.Add Null, "c_TCO_MOV_ABERTO" '7

Set Coligada_Fechamento = Coligada_Critica(ADOConnection, cFields, 3)

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
Public Function Coligada_Critica(ByVal cCONECCAO As Connection, _
                                 ByVal cFields As Collection, _
                                 ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction > 1 Then
   'Critica dos Campos dos Tipos de Formularios
    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
       Err.Raise 50000, Err.Source, "Código Inválido"
    End If
    If Len(Trim(cFields(3))) = 0 Then
       Err.Raise 50000, Err.Source, "Desconto Inválido"
    End If
    If Len(Trim(cFields(3))) = 0 Then
       Err.Raise 50000, Err.Source, "Desconto Inválido"
    End If
    If Len(Trim(cFields(5))) = 0 Then
       Err.Raise 50000, Err.Source, "Verba Inválido"
    End If
    If Len(Trim(cFields(6))) = 0 Then
       Err.Raise 50000, Err.Source, "Ano/mes Inválido"
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
            Case 2, 3, 4, 5
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
                Case 2, 4, 5, 6, 7
                    sSQL = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                
                ' Altera campos tipo data
                'Case 33
                '    sSql = "CONVERT(VARCHAR, '" & Format(cFields(nx), "yyyymmdd") & "', 103), 103)"
                '    'sName = cFields(nx).Name
                '    cFields.Remove nx
                '    cFields.Add sSQL, , , nx - 1
                ' Altera campos tipo data da alteração
                Case 0
                    Sql = "CONVERT(VARCHAR, '" & Format(cFields(nx), "yyyymmdd HH:MM:SS") & "', 108), 108)"
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
                Case 3
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

Set Coligada_Critica = Coligada_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE DADOS DO Feriado
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Coligada_dados(ByVal cCONECCAO As Connection, _
                               ByVal cFields As Collection, _
                               ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String

Select Case nAction
    Case 1 'Consultar Tipos de Formularios

    sSQL = "SELECT " & _
           "CMU_COLIGADA.TCO_CODIGO, " & _
           "CMU_COLIGADA.TCO_DESCRICAO, " & _
           "CMU_COLIGADA.TCO_DESCONTO, " & _
           "CMU_COLIGADA.TCO_DT_MOD, " & _
           "CMU_COLIGADA.TCO_VERBA, " & _
           "CMU_COLIGADA.TCO_ANO_MES_PROC, " & _
           "CMU_COLIGADA.TCO_MOV_ABERTO "
           
    sSQL = sSQL & _
           " FROM CMU_COLIGADA"

    If cFields.Count > 0 Then
       If Len(Trim(cFields(1))) > 0 Then
          sSQL = sSQL & " WHERE TCO_CODIGO = " & cFields(1) & ""
       Else
          sSQL = sSQL & " ORDER BY TCO_CODIGO"
       End If
    Else
       sSQL = sSQL & " ORDER BY TCO_CODIGO"
    End If
    
    Case 2 'Incluir Tipos de Formularios
     sDate = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"
'                    "TCO_CODIGO, "
    sSQL = "INSERT INTO CMU_COLIGADA(" & _
                    " TCO_DESCRICAO ," & _
                    " TCO_DESCONTO ," & _
                    " TCO_DT_MOD ," & _
                    " TCO_VERBA ," & _
                    " TCO_ANO_MES_PROC ," & _
                    " TCO_MOV_ABERTO "
    sSQL = sSQL & ") VALUES (" & _
                    cFields(1) & ", " & _
                    cFields(2) & ", " & _
                    cFields(3) & ", " & _
                    cFields(4) & ", " & _
                    cFields(5) & "," & _
                    cFields(6) & "," & _
                    cFields(7) & ")"

    Case 3 'Alterar Tipos de Formularios
    
     sDate = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"
'                 "  CMU_COLIGADA.TCO_CODIGO = " & cFields(1)
'                    "  TCO_DESCRICAO = " & cFields(2) & _
'                    " ,TCO_DT_MOD = " & cFields(4)

    sSQL = "UPDATE CMU_COLIGADA SET" & _
                    "  TCO_DESCONTO = " & cFields(3) & _
                    " ,TCO_VERBA = " & cFields(5) & _
                    " ,TCO_ANO_MES_PROC = " & cFields(6) & _
                    " ,TCO_MOV_ABERTO = " & cFields(7) & _
                    " WHERE TCO_CODIGO = " & cFields(1)
                    
'cFields.Add c_TCO_CODIGO, "c_TCO_CODIGO" '1
'cFields.Add c_TCO_DESCRICAO, "c_TCO_DESCRICAO" '2
'cFields.Add c_TCO_DESCONTO, "c_TCO_DESCONTO" '3
'cFields.Add c_TCO_DT_MOD, "c_TCO_DT_MOD" '4
'cFields.Add c_TCO_VERBA, "c_TCO_VERBA" '5
'cFields.Add c_TCO_ANO_MES_PROC, "c_TCO_ANO_MES_PROC" '6
'cFields.Add c_TCO_MOV_ABERTO, "c_TCO_MOV_ABERTO" '7
    
    Case 4 'Excluir Tipos de Formularios
    sSQL = "DELETE  FROM CMU_COLIGADA" & _
                 "  WHERE CMU_COLIGADA.TCO_CODIGO = " & cFields(1) & _
                 " AND CMU_COLIGADA.FOR_DTA = " & cFields(8)

    Case 4 'Excluir Tipos de Formularios
    sSQL = "DELETE  FROM CMU_COLIGADA" & _
                 "  WHERE CMU_COLIGADA.TCO_CODIGO = " & cFields(1) & _
                 " AND CMU_COLIGADA.FOR_DTA = " & cFields(8)

End Select

Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch

rs.Open sSQL, cCONECCAO
'Gravar registro de Log (INICIO)
'If nAction > 1 And nAction < 10 Then
'    Dim cFields1 As Collection
'    Set cFields1 = New Collection
'    cFields1.Add Replace(cFields(7), "'", "")
'    cFields1.Add "ESC_TAB_TIPO_FORMULARIO"
'    cFields1.Add nAction
'    cFields1.Add sSql
'    Call CCTempneLog.LOG_DADOS(cCONECCAO, cFields1, 0)
'    Set cFields1 = Nothing
'End If
'Gravar registro de Log (FIM)

'If nAction = 2 Then  'Receber o ultimo registro que acabou dee Incluir
'   sSql = "SELECT MAX(CMU_COLIGADA.TCO_CODIGO) AS ultimoDeTCO_CODIGO" & _
'                 " FROM CMU_COLIGADA;"
'   cCONECCAO.CursorLocation = adUseClientBatch
'   rs.Open sSql, cCONECCAO
'End If


'cCONECCAO.CommitTrans: cCONECCAO.BeginTrans
Set Coligada_dados = rs
Set rs = Nothing
Set rsSeq = Nothing

Exit Function

Erro:
cCONECCAO.RollbackTrans: cCONECCAO.BeginTrans
''cCONECCAO.Close
Set rs = Nothing
'Set cCONECCAO = Nothing
If Err.Number = -2147467259 Then
   Err.Description = "Erro no Banco, Sem permissão de Acesso no Momento"
End If
Err.Raise Err.Number, Err.Source, Err.Description
End Function



