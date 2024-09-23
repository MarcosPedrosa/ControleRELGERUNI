VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'**************** Ações CAD_LOG
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS DO CAD_LOG
' ***********************************************************
' ***********************************************************
' ***********************************************************

Public Function LOG_CONSULTAR(ByVal sBanco As String, _
                              ByVal c_DataINI As String, _
                              ByVal c_DataFIM As String, _
                              ByVal c_Classf As String, _
                              ByVal c_Acao As String, _
                              ByVal c_Usur As String) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_DataINI, "c_DataINI"
cFields.Add c_DataFIM, "c_DataFIM"
cFields.Add c_Classf, "c_Classf"
cFields.Add c_Acao, "c_Acao"
cFields.Add c_Usur, "c_Usur"

Set LOG_CONSULTAR = LOG_DADOS(ADOConnection, cFields, 1)

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
Public Function LOG_INCLUSAO(ByVal cCONECCAO As Connection, _
                             ByVal sLOG_DATA As String, _
                             ByVal sLOG_HORA As String, _
                             ByVal sLOG_USU As String, _
                             ByVal sLOG_TABELA As String, _
                             ByVal sLOG_ACAO As String, _
                             ByVal sLOG_SQL As String, _
                             ByVal sLOG_OBSERVACAO)
                          
Dim cFields As Collection

Set cFields = New Collection

cFields.Add "'" & Format(sLOG_DATA, "yyyymmdd") & "'", "sLOG_DATA"
cFields.Add "'" & Format(sLOG_HORA, "hh:mm:ss") & "'", "sLOG_HORA"
cFields.Add "'" & sLOG_USU & "'", "sLOG_USU"
cFields.Add "'" & sLOG_TABELA & "'", "sLOG_TABELA"
cFields.Add "'" & sLOG_ACAO & "'", "sLOG_ACAO"
cFields.Add "'" & Replace(sLOG_SQL, "'", "") & "'", "sLOG_SQL"
cFields.Add "'" & Replace(sLOG_OBSERVACAO, "'", "") & "'", "sLOG_OBSERVACAO"

Set LOG_INCLUSAO = LOG_DADOS(cCONECCAO, cFields, 2)

Set cFields = Nothing

Exit Function

Erro:
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function LOG_CONSULTAR_Ult_Abertura_Inv(ByVal sBanco As String) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set LOG_CONSULTAR_Ult_Abertura_Inv = LOG_DADOS(ADOConnection, cFields, 3)

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
Public Function LOG_CONSULTAR_Acoes(ByVal sBanco As String) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set LOG_CONSULTAR_Acoes = LOG_DADOS(ADOConnection, cFields, 4)

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
Public Function LOG_CONSULTAR_Acoes_Usuario(ByVal sBanco As String) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set LOG_CONSULTAR_Acoes_Usuario = LOG_DADOS(ADOConnection, cFields, 5)

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

Public Function LOG_DADOS(ByVal cCONECCAO As Connection, _
                          ByVal cFields As Collection, _
                          ByVal nAction As Integer) As ADODB.Recordset

Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String
Dim bWhere As Boolean

On Error GoTo Erro

Select Case nAction
    Case 1 'Consultar CURSO
    bWhere = False
    sSQL = "SELECT LOG_DATA, " & _
                  "LOG_HORA, " & _
                  "LOG_USU, " & _
                  "LOG_OBSERVACAO, " & _
                  "LOG_TABELA, " & _
                  "LOG_ACAO, " & _
                  "LOG_SQL " & _
                  "FROM EXP_TAB_LOG_SISTEMA "
                  
    
    If Len(Trim(cFields(1))) > 0 Then
       sSQL = sSQL & " WHERE LOG_DATA >= '" & cFields(1) & "'": bWhere = True '1-data inicial do periodo
    End If
    
    If Len(Trim(cFields(2))) > 0 Then
       If bWhere Then
          sSQL = sSQL & " AND  LOG_DATA <= '" & cFields(2) & "'": bWhere = True '2-data final do periodo
       Else
          sSQL = sSQL & " WHERE  LOG_DATA <= '" & cFields(2) & "'" '2-data final do periodo
       End If
    End If
    
    If Len(Trim(cFields(4))) > 0 Then
       If bWhere Then
          sSQL = sSQL & " AND  LOG_ACAO = '" & cFields(4) & "'": bWhere = True '4-ACAO
       Else
          sSQL = sSQL & " WHERE  LOG_ACAO '= " & cFields(4) & "'" '4-ACAO
       End If
    End If

    If Len(Trim(cFields(5))) > 0 Then
       If bWhere Then
          sSQL = sSQL & " AND  LOG_USU = '" & cFields(5) & "'": bWhere = True '5-USUARIO
       Else
          sSQL = sSQL & " WHERE  LOG_USU = '" & cFields(5) '5-USUARIO
       End If
    End If
    
    
    If cFields(3) = "1" Then
          sSQL = sSQL & " ORDER BY LOG_USU,LOG_DATA,LOG_HORA "
    Else
          sSQL = sSQL & " ORDER BY LOG_DATA,LOG_HORA "
    End If
    
    
'******************************************************************************************************
'***************    incluir os dados de log das acoes dos usuarios
'******************************************************************************************************

    Case 2
    
    sSQL = "INSERT INTO EXP_TAB_LOG_SISTEMA (" & _
                    "LOG_DATA, " & _
                    "LOG_HORA, " & _
                    "LOG_USU, " & _
                    "LOG_TABELA, " & _
                    "LOG_ACAO, " & _
                    "LOG_SQL, " & _
                    "LOG_OBSERVACAO "
                    
    sSQL = sSQL & ") VALUES (" & _
                    cFields(1) & ", " & _
                    cFields(2) & ", " & _
                    cFields(3) & ", " & _
                    cFields(4) & ", " & _
                    cFields(5) & ", " & _
                    cFields(6) & ", " & _
                    cFields(7) & ")"
                    
'******************************************************************************************************
'***************   Consultar ultima abertura do inventário
'******************************************************************************************************
    Case 3
    
    sSQL = "SELECT TOP 1 LOG_DATA, " & _
                  "LOG_HORA, " & _
                  "LOG_USU, " & _
                  "LOG_OBSERVACAO, " & _
                  "LOG_TABELA, " & _
                  "LOG_ACAO, " & _
                  "LOG_SQL " & _
                  "FROM EXP_TAB_LOG_SISTEMA "
                  
    sSQL = sSQL & " WHERE LOG_ACAO = 'ABERTURA INVENTARIO'"
    sSQL = sSQL & " ORDER BY LOG_DATA DESC"
'******************************************************************************************************
'***************   'Consultar as acoes registradas
'******************************************************************************************************
    Case 4
    
    sSQL = "SELECT DISTINCT LOG_ACAO " & _
                  "FROM EXP_TAB_LOG_SISTEMA "
'******************************************************************************************************
'***************   'Consultar as acoes registradas
'******************************************************************************************************
    Case 5
    
    sSQL = "SELECT DISTINCT LOG_USU " & _
                  "FROM EXP_TAB_LOG_SISTEMA "
                  
End Select
    
Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch

rs.Open sSQL, cCONECCAO

Set LOG_DADOS = rs
Set rs = Nothing

Exit Function

Erro:
Set rs = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function

