VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neExpedicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS DO EXPEDICAO
' ***********************************************************
' ***********************************************************
' ***********************************************************

Public Function EXPEDICAO_Consultar_Cliente(ByVal sBanco As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set EXPEDICAO_Consultar_Cliente = EXPEDICAO_Critica(ADOConnection, cFields, 1)

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
Public Function EXPEDICAO_Consultar_Comp_Pallet(ByVal sBanco As String, _
                                                ByVal sDataINI As String, _
                                                ByVal sDataFIM As String, _
                                                ByVal sSEQ1 As String, _
                                                ByVal sSEQ2 As String, _
                                                ByVal sStatus As String, _
                                                ByVal sCliente As String, _
                                                ByVal sivent As String) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sDataINI, "sDataINI"
cFields.Add sDataFIM, "sDataFIM"
cFields.Add sSEQ1, "sSEQ1"
cFields.Add sSEQ2, "sSEQ2"
cFields.Add sStatus, "sStatus"
cFields.Add sCliente, "sCliente"
cFields.Add sivent, "sivent"

Set EXPEDICAO_Consultar_Comp_Pallet = EXPEDICAO_Critica(ADOConnection, cFields, 2)

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
Public Function EXPEDICAO_Consultar_Um_Pallet(ByVal sBanco As String, _
                                              ByVal spallet As String) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add spallet, "spallet"

Set EXPEDICAO_Consultar_Um_Pallet = EXPEDICAO_Critica(ADOConnection, cFields, 3)

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
Public Function EXPEDICAO_Consultar_Peca_Lote(ByVal sBanco As String, _
                                              ByVal sPeca As String, _
                                              ByVal slote As String, _
                                              ByVal sStatus As String, _
                                              ByVal sCliente As String, _
                                              ByVal sinvent As String, _
                                              ByVal cfieldsP As Collection) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim nx As Integer

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sPeca, "sPeca"
cFields.Add slote, "slote"
cFields.Add sStatus, "sStatus"
cFields.Add sCliente, "sCliente"
cFields.Add sinvent, "sInvent"

If cfieldsP.Count > 0 Then
   For nx = 1 To cfieldsP.Count
       cFields.Add cfieldsP(nx), (Str(nx))
   Next
End If

Set EXPEDICAO_Consultar_Peca_Lote = EXPEDICAO_Critica(ADOConnection, cFields, 4)

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
Public Function EXPEDICAO_Consultar_Pallet_Peca_Lote(ByVal sBanco As String, _
                                                     ByVal sPeca As String, _
                                                     ByVal slote As String, _
                                                     ByVal spallet As String, _
                                                     ByVal sCliente As String, _
                                                     ByVal cfieldsP As Collection) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add spallet, "spallet"
cFields.Add sPeca, "sPeca"
cFields.Add slote, "slote"
cFields.Add sCliente, "sCliente"

If cfieldsP.Count > 0 Then
   For nx = 1 To cfieldsP.Count
       cFields.Add cfieldsP(nx), (Str(nx))
   Next
End If

Set EXPEDICAO_Consultar_Pallet_Peca_Lote = EXPEDICAO_Critica(ADOConnection, cFields, 5)

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
Public Function EXPEDICAO_Consultar_caixa(ByVal sBanco As String, _
                                          ByVal sCx1 As String, _
                                          ByVal sCx2 As String, _
                                          ByVal sStatus As String, _
                                          ByVal sCliente As String, _
                                          ByVal sinvent As String, _
                                          ByVal cFieldsC As Collection) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sCx1, "sCx1"
cFields.Add sCx2, "sCx2"
cFields.Add sStatus, "sStatus"
cFields.Add sCliente, "sCliente"
cFields.Add sinvent, "sinvent"

If cFieldsC.Count > 0 Then
   For nx = 1 To cFieldsC.Count
       cFields.Add cFieldsC(nx), (Str(nx))
   Next
End If

Set EXPEDICAO_Consultar_caixa = EXPEDICAO_Critica(ADOConnection, cFields, 6)

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
Public Function EXPEDICAO_Consultar_Pallet_Caixa(ByVal sBanco As String, _
                                                 ByVal sCx1 As String, _
                                                 ByVal sCx2 As String, _
                                                 ByVal spallet As String, _
                                                 ByVal sCliente As String, _
                                                 ByVal cFieldsC As Collection) As ADODB.Recordset


On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add spallet, "spallet"
cFields.Add sCx1, "sCx1"
cFields.Add sCx2, "sCx2"
cFields.Add sCliente, "sCliente"

If cFieldsC.Count > 0 Then
   For nx = 1 To cFieldsC.Count
       cFields.Add cFieldsC(nx), (Str(nx))
   Next
End If


Set EXPEDICAO_Consultar_Pallet_Caixa = EXPEDICAO_Critica(ADOConnection, cFields, 7)

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
Public Function EXPEDICAO_Consultar_Intervalo_Caixa(ByVal sBanco As String, _
                                                    ByVal sCx1 As String, _
                                                    ByVal sCx2 As String, _
                                                    ByVal sStatus As String) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sCx1, "sCx1"
cFields.Add sCx2, "sCx2"
cFields.Add sStatus, "sStatus"

Set EXPEDICAO_Consultar_Intervalo_Caixa = EXPEDICAO_Critica(ADOConnection, cFields, 8)

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
Public Function EXPEDICAO_Consultar_Seq_Placa(ByVal sBanco As String, _
                                              ByVal sDataINI As String, _
                                              ByVal sDataFIM As String, _
                                              ByVal sSEQ1 As String, _
                                              ByVal sSEQ2 As String, _
                                              ByVal sStatus As String, _
                                              ByVal sCliente As String, _
                                              ByVal sPlaca1 As String, _
                                              ByVal sPlaca2 As String) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sDataINI, "sDataINI"
cFields.Add sDataFIM, "sDataFIM"
cFields.Add sSEQ1, "sSEQ1"
cFields.Add sSEQ2, "sSEQ2"
cFields.Add sStatus, "sStatus"
cFields.Add sCliente, "sCliente"
cFields.Add sPlaca1, "sPlaca1"
cFields.Add sPlaca2, "sPlaca2"

Set EXPEDICAO_Consultar_Seq_Placa = EXPEDICAO_Critica(ADOConnection, cFields, 9)

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
Public Function EXPEDICAO_Consultar_Pallet_Seq_Caixa(ByVal sBanco As String, _
                                                     ByVal sCx1 As String, _
                                                     ByVal sCx2 As String, _
                                                     ByVal spallet As String, _
                                                     ByVal sStatus As String, _
                                                     ByVal sCliente As String) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add spallet, "spallet"
cFields.Add sCx1, "sCx1"
cFields.Add sCx2, "sCx2"
cFields.Add sCliente, "sCliente"
cFields.Add sStatus, "sStatus"

Set EXPEDICAO_Consultar_Pallet_Seq_Caixa = EXPEDICAO_Critica(ADOConnection, cFields, 10)

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
Public Function EXPEDICAO_Atualiza_Inventario(ByVal sBanco As String, _
                                              ByVal sID_ETIQUETA As String, _
                                              ByVal sID_BORDERO As String) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sID_ETIQUETA, "sID_ETIQUETA"
cFields.Add sID_BORDERO, "sID_BORDERO"

Call EXPEDICAO_Critica(ADOConnection, cFields, 11)

ADOConnection.CommitTrans
'Set ADOConnection = Nothing
'Set cConect = Nothing
'Set cFields = Nothing

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
Public Function EXPEDICAO_Consultar_Comp_NF(ByVal sBanco As String, _
                                            ByVal sDataINI As String, _
                                            ByVal sDataFIM As String, _
                                            ByVal sSEQ1 As String, _
                                            ByVal sSEQ2 As String, _
                                            ByVal sStatus As String, _
                                            ByVal sCliente As String, _
                                            ByVal sNF1 As String, _
                                            ByVal sNF2 As String) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sDataINI, "sDataINI"
cFields.Add sDataFIM, "sDataFIM"
cFields.Add sSEQ1, "sSEQ1"
cFields.Add sSEQ2, "sSEQ2"
cFields.Add sStatus, "sStatus"
cFields.Add sCliente, "sCliente"
cFields.Add sNF1, "sNF1"
cFields.Add sNF2, "sNF2"

Set EXPEDICAO_Consultar_Comp_NF = EXPEDICAO_Critica(ADOConnection, cFields, 12)

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
Public Function EXPEDICAO_Consultar_Pallet_Seq_NF(ByVal sBanco As String, _
                                                  ByVal sNF1 As String, _
                                                  ByVal sNF2 As String, _
                                                  ByVal spallet As String, _
                                                  ByVal sStatus As String, _
                                                  ByVal sCliente As String) As ADODB.Recordset

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add spallet, "spallet"
cFields.Add sNF1, "sNF1"
cFields.Add sNF2, "sNF2"
cFields.Add sCliente, "sCliente"
cFields.Add sStatus, "sStatus"

Set EXPEDICAO_Consultar_Pallet_Seq_NF = EXPEDICAO_Critica(ADOConnection, cFields, 13)

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
Public Function EXPEDICAO_Consultar_Todas_Pecas(ByVal sBanco As String, _
                                                ByVal sCliente As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sCliente, "scliente"

Set EXPEDICAO_Consultar_Todas_Pecas = EXPEDICAO_Critica(ADOConnection, cFields, 14)

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
Public Function EXPEDICAO_Consultar_Uma_Caixa(ByVal sBanco As String, _
                                              ByVal scaixa As String, _
                                              ByVal sCliente As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add scaixa, "scaixa"
cFields.Add sCliente, "sCliente"

Set EXPEDICAO_Consultar_Uma_Caixa = EXPEDICAO_Critica(ADOConnection, cFields, 15)

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

Public Function EXPEDICAO_Critica(ByVal cCONECCAO As Connection, _
                                ByVal cFields As Collection, _
                                ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String


If nAction > 100 Then
   'Critica dos Campos dos USUARIO
'    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
'       Err.Raise 50000, Err.Source, "Código do USUARIO Inválido"
'    End If
'
'    If Len(Trim(cFields(2))) = 0 Then
'       Err.Raise 50000, Err.Source, "Nome do USUARIO Inválido"
'    End If
'
'    If Len(Trim(cFields(3))) = 0 Then
'       Err.Raise 50000, Err.Source, "Codigo do USUARIO Inválido"
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
                Case 2, 3, 4
                 If Len(Trim(cFields(nx))) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "NULL", , , nx - 1
                 End If
                
                'Verifica campos numéricos NULL
                Case 0
                    If Val(cFields(nx)) = 0 Then
                        'sName = cFields(nx).Name
                        cFields.Remove nx
                        cFields.Add "NULL", , , nx - 1
                    End If
            End Select
        End If
        
        If cFields(nx) <> "NULL" Then
            Select Case nx
                'Altera campos alfanuméricos
                Case 2, 3, 4
                    sSQL = "'" & Replace(cFields(nx), "'", "`") & "'"
'                    sSQL = "'" & cFields(nx) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                
                ' Altera campos tipo data
                Case 0
                    sSQL = "CONVERT(VARCHAR, '" & Format(cFields(nx), "yyyymmdd") & "', 103), 103)"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                 'Altera campos tipo data da alteração
                Case 5
                    sSQL = "CONVERT(VARCHAR, '" & Format(cFields(nx), "yyyymmdd HH:MM:SS") & "', 108), 108)"
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
                Case 0
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

Set EXPEDICAO_Critica = EXPEDICAO_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE DADOS DO USUARIO
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function EXPEDICAO_dados(ByVal cCONECCAO As Connection, _
                                ByVal cFields As Collection, _
                                ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String
Dim nseq As Integer
Dim rsSeq As ADODB.Recordset
Dim bWhere As Boolean

Select Case nAction
'****************************************************************************************************
'********** consultar os clientes na tabela mov_etiq
'****************************************************************************************************
    Case 1 'Consultar clientes
    
    sSQL = "SELECT " & _
           "ID_CLIENTE, " & _
           "CLIENTE "
    sSQL = sSQL & _
           "FROM MOV_ETIQ " & _
           "GROUP BY Cliente, ID_CLIENTE "
           
'****************************************************************************************************
'********** Consultar ETIQUETAS DE ACORDO COM O FILTRO ABAIXO
'****************************************************************************************************
    Case 2
    
    sSQL = "SELECT " & _
           "PALLET, " & _
           "ID_CLIENTE, " & _
           "CASE CONF_PALLET WHEN 'E' THEN 'Embarcou' ELSE " & _
           "CASE FATUR WHEN 'F' THEN 'Faturado' ELSE 'Nao fatur' " & _
           "END END AS CONF_PALLET, " & _
           "(SELECT TOP 1 CLIENTE " & _
           "FROM MOV_ETIQ E " & _
           "WHERE E.ID_CLIENTE = ETIQUETA.ID_CLIENTE) AS CLIENTE, " & _
           "COUNT(*) AS CAIXA," & _
           "SUM(QTDE) AS QTDE "
    sSQL = sSQL & _
           "FROM ETIQUETA "
    sSQL = sSQL & _
           "WHERE  PALLET IS NOT NULL " & _
           "AND  PALLET <> '' "
    
    If Len(Trim(cFields(1))) > 0 Then
       sSQL = sSQL & _
              "AND  (SUBSTRING(pallet, 5, 4) + SUBSTRING(pallet, 3, 2) + SUBSTRING(pallet, 1, 2)) >= '" & cFields(1) & "' "
       sSQL = sSQL & _
              "AND SUBSTRING(pallet, 10, 3) >= '" & cFields(3) & "' "
    End If
    
    If Len(Trim(cFields(2))) > 0 Then
       sSQL = sSQL & _
              "AND  (SUBSTRING(pallet, 5, 4) + SUBSTRING(pallet, 3, 2) + SUBSTRING(pallet, 1, 2)) <= '" & cFields(2) & "' "
       sSQL = sSQL & _
              "AND SUBSTRING(pallet, 10, 3) <= '" & cFields(4) & "' "
    End If

    If cFields(5) = "1" Then 'FATURADO
       sSQL = sSQL & _
              "AND  CONF_PALLET = 'F' " & _
              "AND  FATUR <> '' "
    ElseIf cFields(5) = "2" Then 'EMBARCOU
       sSQL = sSQL & _
              "AND  CONF_PALLET = 'E' "
    ElseIf cFields(5) = "3" Then 'NAO FATURADO
       sSQL = sSQL & _
              "AND  CONF_PALLET = 'P' " & _
              "AND  FATUR = '' "
    ElseIf cFields(5) = "4" Then 'NAO EMBARCOU (NA FABRICA)
       sSQL = sSQL & _
              "AND  CONF_PALLET = 'P' "
    End If

    If Len(Trim(cFields(6))) > 0 Then
       sSQL = sSQL & _
              "AND  ID_CLIENTE = '" & cFields(6) & "' "
    End If

'Me.CBO_INVENTARIO.AddItem "Todos"
'Me.CBO_INVENTARIO.AddItem "Não inventariado"
'Me.CBO_INVENTARIO.AddItem "Com contagem"
'Me.CBO_INVENTARIO.AddItem "1a Contagem"
'Me.CBO_INVENTARIO.AddItem "2a Contagem"
'Me.CBO_INVENTARIO.AddItem "3a Contagem"
'Me.CBO_INVENTARIO.ListIndex = 0

    If cFields(7) = "0" Then 'TODOS
       sSQL = sSQL & _
              " "
    ElseIf cFields(7) = "1" Then 'Não inventariado
       sSQL = sSQL & _
              "AND  ID_BORDERO = '' "
    ElseIf cFields(7) = "2" Then 'Com contagem
       sSQL = sSQL & _
              "AND  ID_BORDERO <> '' "
    Else 'NUMERO DA CONTAGEM
       sSQL = sSQL & _
              "AND  ID_BORDERO = '" & Val(cFields(7)) - 2 & "' "
    End If


    sSQL = sSQL & _
           "GROUP BY PALLET,ID_CLIENTE,CONF_PALLET,FATUR "
           
'****************************************************************************************************
'********** Consultar ETIQUETAS PELO NUMERO DO PALLET
'****************************************************************************************************
    Case 3
    
    sSQL = "SELECT " & _
           "ID_ETIQUETA AS NUM_CAIXA, " & _
           "CASE TIPO_PADRAO_CX WHEN 1 THEN 'IK10' ELSE " & _
           "CASE TIPO_PADRAO_CX WHEN 2 THEN 'IK33' ELSE 'PAPELAO' " & _
           "END END AS TIPO_CAIXA, " & _
           "ID_PECA AS PECA," & _
           "LOTE," & _
           "QTDE AS QTDE_NA_CAIXA," & _
           "XBLNR AS NF_VENDA, " & _
           "REMESSA AS ORDEM_VENDA," & _
           "SEQUENCIA_PLACA AS SEQUENCIA, " & _
           "PLACA "
    sSQL = sSQL & _
           "FROM ETIQUETA "
    sSQL = sSQL & _
           "WHERE  PALLET = '" & cFields(1) & "' "
              
    sSQL = sSQL & _
           "ORDER BY ID_PECA "
           

'****************************************************************************************************
'********** Consultar ETIQUETAS DE ACORDO COM O FILTRO ABAIXO (pecas e lote)
'****************************************************************************************************
    Case 4
    
    bWhere = False
    
    sSQL = "SELECT " & _
           "PALLET, " & _
           "ID_CLIENTE, " & _
           "CASE CONF_PALLET WHEN 'E' THEN 'Embarcou' ELSE " & _
           "CASE FATUR WHEN 'F' THEN 'Faturado' ELSE 'Nao fatur' " & _
           "END END AS CONF_PALLET, " & _
           "(SELECT TOP 1 CLIENTE " & _
           "FROM MOV_ETIQ E " & _
           "WHERE E.ID_CLIENTE = ETIQUETA.ID_CLIENTE) AS CLIENTE, " & _
           "COUNT(*) AS CAIXA," & _
           "SUM(QTDE) AS QTDE "
    sSQL = sSQL & _
           "FROM ETIQUETA "
'    sSQL = sSQL & _
'           "WHERE  PALLET IS NOT NULL " & _
'           "AND  PALLET <> '' "
    
    If Len(Trim(cFields(1))) > 0 Then
       sSQL = sSQL & _
              "WHERE ID_PECA = '" & cFields(1) & "' ": bWhere = True
    End If

    If Len(Trim(cFields(2))) > 0 Then
       If bWhere Then
          sSQL = sSQL & " AND "
       Else
          sSQL = sSQL & " WHERE "
       End If
       sSQL = sSQL & _
              "LOTE = '" & cFields(2) & "' ": bWhere = True
    End If

    If bWhere Then
       sSQL = sSQL & " AND "
    Else
       sSQL = sSQL & " WHERE "
    End If
    If cFields(3) = "1" Then 'FATURADO
       sSQL = sSQL & _
              "CONF_PALLET = 'F' AND " & _
              "FATUR <> '' ": bWhere = True
    ElseIf cFields(3) = "2" Then 'EMBARCOU
       sSQL = sSQL & _
              "CONF_PALLET = 'E' ": bWhere = True
    ElseIf cFields(3) = "3" Then 'NAO FATURADO
       sSQL = sSQL & _
              "(CONF_PALLET = 'P' OR CONF_PALLET = '') AND " & _
              "FATUR = '' ": bWhere = True
    ElseIf cFields(3) = "4" Then 'NAO EMBARCOU (NA FABRICA)
       sSQL = sSQL & _
              "(CONF_PALLET = 'P' OR CONF_PALLET = '') ": bWhere = True
    End If

    If Len(Trim(cFields(4))) > 0 Then
       sSQL = sSQL & _
              "AND  ID_CLIENTE = '" & cFields(4) & "' "
    End If

'Me.CBO_INVENTARIO.AddItem "Todos"
'Me.CBO_INVENTARIO.AddItem "Não inventariado"
'Me.CBO_INVENTARIO.AddItem "Com contagem"
'Me.CBO_INVENTARIO.AddItem "1a Contagem"
'Me.CBO_INVENTARIO.AddItem "2a Contagem"
'Me.CBO_INVENTARIO.AddItem "3a Contagem"
'Me.CBO_INVENTARIO.ListIndex = 0

    If cFields(5) = "0" Then 'TODOS
       sSQL = sSQL & _
              " "
    ElseIf cFields(5) = "1" Then 'Não inventariado
       sSQL = sSQL & _
              "AND  ID_BORDERO = '' "
    ElseIf cFields(5) = "2" Then 'Com contagem
       sSQL = sSQL & _
              "AND  ID_BORDERO <> '' "
    Else 'NUMERO DA CONTAGEM
       sSQL = sSQL & _
              "AND  ID_BORDERO = '" & Val(cFields(5)) - 2 & "' "
    End If
    
Rem verificar se ha pecas escolhidas
    If cFields.Count > 5 Then
       sSQL = sSQL & _
              "AND  (ID_PECA = '" & cFields(6) & "' "
       For nx = 7 To cFields.Count
           sSQL = sSQL & _
                  " OR ID_PECA = '" & cFields(nx) & "' "
       Next
       sSQL = sSQL & ") "
    End If
    
    sSQL = sSQL & _
           " GROUP BY PALLET,ID_CLIENTE,CONF_PALLET,FATUR "

'****************************************************************************************************
'********** Consultar ETIQUETAS PELO NUMERO DO PALLET/PECA E LOTE
'****************************************************************************************************
    Case 5
    
    sSQL = "SELECT " & _
           "ID_ETIQUETA AS NUM_CAIXA, " & _
           "CASE TIPO_PADRAO_CX WHEN 1 THEN 'IK10' ELSE " & _
           "CASE TIPO_PADRAO_CX WHEN 2 THEN 'IK33' ELSE 'PAPELAO' " & _
           "END END AS TIPO_CAIXA, " & _
           "ID_PECA AS PECA," & _
           "LOTE," & _
           "QTDE AS QTDE_NA_CAIXA," & _
           "XBLNR AS NF_VENDA, " & _
           "REMESSA AS ORDEM_VENDA," & _
           "SEQUENCIA_PLACA AS SEQUENCIA, " & _
           "PLACA "
    sSQL = sSQL & _
           "FROM ETIQUETA "
           
    sSQL = sSQL & _
           "WHERE  ID_CLIENTE = '" & cFields(4) & "' "
    
    sSQL = sSQL & _
           "AND  PALLET = '" & cFields(1) & "' "
              
    If Len(Trim(cFields(2))) > 0 Then
       sSQL = sSQL & _
              "AND ID_PECA = '" & cFields(2) & "' "
    End If

    If Len(Trim(cFields(3))) > 0 Then
       sSQL = sSQL & _
              "AND LOTE = '" & cFields(3) & "' "
    End If
    
Rem verificar se ha pecas escolhidas
    If cFields.Count > 4 Then
       sSQL = sSQL & _
              "AND  (ID_PECA = '" & cFields(5) & "' "
       For nx = 6 To cFields.Count
           sSQL = sSQL & _
                  " OR ID_PECA = '" & cFields(nx) & "' "
       Next
       sSQL = sSQL & ") "
    End If
    
    sSQL = sSQL & _
           " ORDER BY ID_PECA "
           
'****************************************************************************************************
'********** Consultar ETIQUETAS DE ACORDO COM O FILTRO ABAIXO (caixa)
'****************************************************************************************************
    Case 6
    
    bWhere = False
    
    sSQL = "SELECT " & _
           "PALLET, " & _
           "ID_CLIENTE, " & _
           "CASE CONF_PALLET WHEN 'E' THEN 'Embarcou' ELSE " & _
           "CASE FATUR WHEN 'F' THEN 'Faturado' ELSE 'Nao fatur' " & _
           "END END AS CONF_PALLET, " & _
           "(SELECT TOP 1 CLIENTE " & _
           "FROM MOV_ETIQ E " & _
           "WHERE E.ID_CLIENTE = ETIQUETA.ID_CLIENTE) AS CLIENTE, " & _
           "COUNT(*) AS CAIXA," & _
           "SUM(QTDE) AS QTDE "
    sSQL = sSQL & _
           "FROM ETIQUETA "
'    sSQL = sSQL & _
'           "WHERE  PALLET IS NOT NULL " & _
'           "AND  PALLET <> '' "
    
    If Len(Trim(cFields(1))) > 0 Then
       sSQL = sSQL & _
              "WHERE ID_ETIQUETA >= '" & cFields(1) & "' ": bWhere = True
    End If

    If Len(Trim(cFields(2))) > 0 Then
       If bWhere Then
          sSQL = sSQL & " AND "
       Else
          sSQL = sSQL & " WHERE "
       End If
       sSQL = sSQL & _
              "ID_ETIQUETA <= '" & cFields(2) & "' ": bWhere = True
    End If

    If bWhere Then
       sSQL = sSQL & " AND "
    Else
       sSQL = sSQL & " WHERE "
    End If
    If cFields(3) = "1" Then 'FATURADO
       sSQL = sSQL & _
              "CONF_PALLET = 'F' AND " & _
              "FATUR <> '' ": bWhere = True
    ElseIf cFields(3) = "2" Then 'EMBARCOU
       sSQL = sSQL & _
              "CONF_PALLET = 'E' ": bWhere = True
    ElseIf cFields(3) = "3" Then 'NAO FATURADO
       sSQL = sSQL & _
              "(CONF_PALLET = 'P' OR CONF_PALLET = '') AND " & _
              "FATUR = '' ": bWhere = True
    ElseIf cFields(3) = "4" Then 'NAO EMBARCOU (NA FABRICA)
       sSQL = sSQL & _
              "(CONF_PALLET = 'P' OR CONF_PALLET = '') ": bWhere = True
    End If

    If Len(Trim(cFields(4))) > 0 Then
       sSQL = sSQL & _
              "AND  ID_CLIENTE = '" & cFields(4) & "' "
    End If


'Me.CBO_INVENTARIO.AddItem "Todos"
'Me.CBO_INVENTARIO.AddItem "Não inventariado"
'Me.CBO_INVENTARIO.AddItem "Com contagem"
'Me.CBO_INVENTARIO.AddItem "1a Contagem"
'Me.CBO_INVENTARIO.AddItem "2a Contagem"
'Me.CBO_INVENTARIO.AddItem "3a Contagem"
'Me.CBO_INVENTARIO.ListIndex = 0

    If cFields(5) = "0" Then 'TODOS
       sSQL = sSQL & _
              " "
    ElseIf cFields(5) = "1" Then 'Não inventariado
       sSQL = sSQL & _
              "AND  ID_BORDERO = '' "
    ElseIf cFields(5) = "2" Then 'Com contagem
       sSQL = sSQL & _
              "AND  ID_BORDERO <> '' "
    Else 'NUMERO DA CONTAGEM
       sSQL = sSQL & _
              "AND  ID_BORDERO = '" & Val(cFields(5)) - 2 & "' "
    End If

Rem verificar se ha caixas escolhidas
    If cFields.Count > 5 Then
       sSQL = sSQL & _
              "AND  (ID_ETIQUETA = '" & cFields(6) & "' "
       For nx = 7 To cFields.Count
           sSQL = sSQL & _
                  " OR ID_ETIQUETA = '" & cFields(nx) & "' "
       Next
       sSQL = sSQL & ") "
    End If

    sSQL = sSQL & _
           "GROUP BY PALLET,ID_CLIENTE,CONF_PALLET,FATUR "
           
'****************************************************************************************************
'********** Consultar ETIQUETAS PELO NUMERO DO CLIENTE/PALLET/INTERVALO DE CAIXAS
'****************************************************************************************************
    Case 7
    
    sSQL = "SELECT " & _
           "ID_ETIQUETA AS NUM_CAIXA, " & _
           "CASE TIPO_PADRAO_CX WHEN 1 THEN 'IK10' ELSE " & _
           "CASE TIPO_PADRAO_CX WHEN 2 THEN 'IK33' ELSE 'PAPELAO' " & _
           "END END AS TIPO_CAIXA, " & _
           "ID_PECA AS PECA," & _
           "LOTE," & _
           "QTDE AS QTDE_NA_CAIXA," & _
           "XBLNR AS NF_VENDA, " & _
           "REMESSA AS ORDEM_VENDA," & _
           "SEQUENCIA_PLACA AS SEQUENCIA, " & _
           "PLACA "
    sSQL = sSQL & _
           "FROM ETIQUETA "
           
    sSQL = sSQL & _
           "WHERE  ID_CLIENTE = '" & cFields(4) & "' "
    
    sSQL = sSQL & _
           "AND  PALLET = '" & cFields(1) & "' "
              
    If Len(Trim(cFields(2))) > 0 Then
       sSQL = sSQL & _
              "AND ID_ETIQUETA >= '" & cFields(2) & "' "
    End If

    If Len(Trim(cFields(3))) > 0 Then
       sSQL = sSQL & _
              "AND ID_ETIQUETA <= '" & cFields(3) & "' "
    End If
    
Rem verificar se ha caixas escolhidas
    If cFields.Count > 4 Then
       sSQL = sSQL & _
              "AND  (ID_ETIQUETA = '" & cFields(5) & "' "
       For nx = 5 To cFields.Count
           sSQL = sSQL & _
                  " OR ID_ETIQUETA = '" & cFields(nx) & "' "
       Next
       sSQL = sSQL & ") "
    End If
    
    sSQL = sSQL & _
           "ORDER BY ID_PECA "
'****************************************************************************************************
'********** Consultar ETIQUETAS PELO INTEVALO DE CAIXA
'****************************************************************************************************
    Case 8
    
    bWhere = False
    
    sSQL = "SELECT " & _
           "ID_ETIQUETA AS NUM_CAIXA, " & _
           "CASE TIPO_PADRAO_CX WHEN 1 THEN 'IK10' ELSE " & _
           "CASE TIPO_PADRAO_CX WHEN 2 THEN 'IK33' ELSE 'PAPELAO' " & _
           "END END AS TIPO_CAIXA, " & _
           "ID_PECA AS PECA," & _
           "LOTE," & _
           "QTDE AS QTDE_NA_CAIXA," & _
           "XBLNR AS NF_VENDA, " & _
           "REMESSA AS ORDEM_VENDA," & _
           "SEQUENCIA_PLACA AS SEQUENCIA, " & _
           "PLACA "
    sSQL = sSQL & _
           "FROM ETIQUETA "
           
    If Len(Trim(cFields(1))) > 0 Then
       sSQL = sSQL & _
              " WHERE ID_ETIQUETA >= '" & cFields(1) & "' ": bWhere = True
    End If
    
    If Len(Trim(cFields(2))) > 0 Then
       If bWhere Then
          sSQL = sSQL & " AND "
       Else
          sSQL = sSQL & " WHERE "
       End If
       sSQL = sSQL & _
              " ID_ETIQUETA <= '" & cFields(2) & "' "
    End If

    If bWhere Then
       sSQL = sSQL & " AND "
    Else
       sSQL = sSQL & " WHERE "
    End If
    
    If cFields(3) = "1" Then 'FATURADO
       sSQL = sSQL & _
              "CONF_PALLET = 'F' AND " & _
              "FATUR <> '' ": bWhere = True
    ElseIf cFields(3) = "2" Then 'EMBARCOU
       sSQL = sSQL & _
              "CONF_PALLET = 'E' ": bWhere = True
    ElseIf cFields(3) = "3" Then 'NAO FATURADO
       sSQL = sSQL & _
              "(CONF_PALLET = 'P' OR CONF_PALLET = '') AND " & _
              "FATUR = '' ": bWhere = True
    ElseIf cFields(3) = "4" Then 'NAO EMBARCOU (NA FABRICA)
       sSQL = sSQL & _
              "(CONF_PALLET = 'P' OR CONF_PALLET = '') ": bWhere = True
    End If
    
    sSQL = sSQL & _
           " ORDER BY ID_ETIQUETA "

'****************************************************************************************************
'********** Consultar ETIQUETAS DE ACORDO COM O FILTRO ABAIXO (intervalo de sequencia das placas)
'****************************************************************************************************
    
    Case 9
    
    sSQL = "SELECT " & _
           "PALLET, " & _
           "ID_CLIENTE, " & _
           "CASE CONF_PALLET WHEN 'E' THEN 'Embarcou' ELSE " & _
           "CASE FATUR WHEN 'F' THEN 'Faturado' ELSE 'Nao fatur' " & _
           "END END AS CONF_PALLET, " & _
           "(SELECT TOP 1 CLIENTE " & _
           "FROM MOV_ETIQ E " & _
           "WHERE E.ID_CLIENTE = ETIQUETA.ID_CLIENTE) AS CLIENTE, " & _
           "COUNT(*) AS CAIXA," & _
           "SUM(QTDE) AS QTDE "
    sSQL = sSQL & _
           "FROM ETIQUETA "
    sSQL = sSQL & _
           "WHERE  PALLET IS NOT NULL " & _
           "AND  PALLET <> '' "
    
    If Len(Trim(cFields(1))) > 0 Then
       sSQL = sSQL & _
              "AND  (SUBSTRING(pallet, 5, 4) + SUBSTRING(pallet, 3, 2) + SUBSTRING(pallet, 1, 2)) >= '" & cFields(1) & "' "
       sSQL = sSQL & _
              "AND SUBSTRING(pallet, 10, 3) >= '" & cFields(3) & "' "
    End If
    
    If Len(Trim(cFields(2))) > 0 Then
       sSQL = sSQL & _
              "AND  (SUBSTRING(pallet, 5, 4) + SUBSTRING(pallet, 3, 2) + SUBSTRING(pallet, 1, 2)) <= '" & cFields(2) & "' "
       sSQL = sSQL & _
              "AND SUBSTRING(pallet, 10, 3) <= '" & cFields(4) & "' "
    End If

    If cFields(5) = "1" Then 'FATURADO
       sSQL = sSQL & _
              "AND  CONF_PALLET = 'F' " & _
              "AND  FATUR <> '' "
    ElseIf cFields(5) = "2" Then 'EMBARCOU
       sSQL = sSQL & _
              "AND  CONF_PALLET = 'E' "
    ElseIf cFields(5) = "3" Then 'NAO FATURADO
       sSQL = sSQL & _
              "AND  CONF_PALLET = 'P' " & _
              "AND  FATUR = '' "
    ElseIf cFields(5) = "4" Then 'NAO EMBARCOU (NA FABRICA)
       sSQL = sSQL & _
              "AND  CONF_PALLET = 'P' "
    End If

    If Len(Trim(cFields(6))) > 0 Then
       sSQL = sSQL & _
              "AND  ID_CLIENTE = '" & cFields(6) & "' "
    End If

    If Len(Trim(cFields(7))) > 0 Then
       sSQL = sSQL & _
              "AND  SEQUENCIA_PLACA >= " & cFields(7)
    End If
    
    If Len(Trim(cFields(8))) > 0 Then
       sSQL = sSQL & _
              "AND  SEQUENCIA_PLACA <= " & cFields(8)
    End If
    
    sSQL = sSQL & _
           " GROUP BY PALLET,ID_CLIENTE,CONF_PALLET,FATUR "
           
'****************************************************************************************************
'********** Consultar ETIQUETAS PELO NUMERO DO CLIENTE/PALLET/INTERVALO DE CAIXAS
'****************************************************************************************************
    Case 10
    
    sSQL = "SELECT " & _
           "ID_ETIQUETA AS NUM_CAIXA, " & _
           "CASE TIPO_PADRAO_CX WHEN 1 THEN 'IK10' ELSE " & _
           "CASE TIPO_PADRAO_CX WHEN 2 THEN 'IK33' ELSE 'PAPELAO' " & _
           "END END AS TIPO_CAIXA, " & _
           "ID_PECA AS PECA," & _
           "LOTE," & _
           "QTDE AS QTDE_NA_CAIXA," & _
           "XBLNR AS NF_VENDA, " & _
           "REMESSA AS ORDEM_VENDA," & _
           "SEQUENCIA_PLACA AS SEQUENCIA, " & _
           "PLACA "
    sSQL = sSQL & _
           "FROM ETIQUETA "
           
    sSQL = sSQL & _
           "WHERE  ID_CLIENTE = '" & cFields(4) & "' "
    
    sSQL = sSQL & _
           "AND  PALLET = '" & cFields(1) & "' "
              
    If Len(Trim(cFields(2))) > 0 Then
       sSQL = sSQL & _
              "AND SEQUENCIA_PLACA >= " & cFields(2)
    End If

    If Len(Trim(cFields(3))) > 0 Then
       sSQL = sSQL & _
              "AND SEQUENCIA_PLACA <= " & cFields(3)
    End If
    
    If cFields(5) = "1" Then 'FATURADO
       sSQL = sSQL & _
              " AND CONF_PALLET = 'F' AND " & _
              "FATUR <> '' ": bWhere = True
    ElseIf cFields(5) = "2" Then 'EMBARCOU
       sSQL = sSQL & _
              " AND CONF_PALLET = 'E' ": bWhere = True
    ElseIf cFields(5) = "3" Then 'NAO FATURADO
       sSQL = sSQL & _
              " AND (CONF_PALLET = 'P' OR CONF_PALLET = '') AND " & _
              "FATUR = '' ": bWhere = True
    ElseIf cFields(5) = "4" Then 'NAO EMBARCOU (NA FABRICA)
       sSQL = sSQL & _
              " AND (CONF_PALLET = 'P' OR CONF_PALLET = '') ": bWhere = True
    End If
    
    sSQL = sSQL & _
           " ORDER BY ID_PECA "
'****************************************************************************************************
'********** ATUALIZAR ETIQUETAS PELO RECEBENDO O NUMERO DA ETIQUETA E O CAMPO BORDERO
'****************************************************************************************************
    Case 11
    
    sSQL = "SELECT COUNT(*) AS CONTADOR FROM ETIQUETA WHERE " & _
           "ID_BORDERO = '" & Trim(cFields(2)) & "'" & _
           " AND ID_ETIQUETA = '" & cFields(1) & "'"
    
    Set rs = New ADODB.Recordset
    cCONECCAO.CursorLocation = adUseClientBatch
    rs.Open sSQL, cCONECCAO
    
    If rs!CONTADOR = 0 Then
       nAtualizadas = nAtualizadas + 1
       sSQL = "UPDATE ETIQUETA SET " & _
              "ID_BORDERO = '" & Trim(cFields(2)) & "'" & _
              " WHERE ID_ETIQUETA = '" & cFields(1) & "'"
    Else
       sSQL = ""
    End If
    
'****************************************************************************************************
'********** Consultar ETIQUETAS DE ACORDO COM O FILTRO ABAIXO (intervalo de sequencia das nf)
'****************************************************************************************************
    
    Case 12
    
    sSQL = "SELECT " & _
           "PALLET, " & _
           "ID_CLIENTE, " & _
           "CASE CONF_PALLET WHEN 'E' THEN 'Embarcou' ELSE " & _
           "CASE FATUR WHEN 'F' THEN 'Faturado' ELSE 'Nao fatur' " & _
           "END END AS CONF_PALLET, " & _
           "(SELECT TOP 1 CLIENTE " & _
           "FROM MOV_ETIQ E " & _
           "WHERE E.ID_CLIENTE = ETIQUETA.ID_CLIENTE) AS CLIENTE, " & _
           "COUNT(*) AS CAIXA," & _
           "SUM(QTDE) AS QTDE "
    sSQL = sSQL & _
           "FROM ETIQUETA "
    sSQL = sSQL & _
           "WHERE  CONF_PALLET <> '' "
    
    If Len(Trim(cFields(1))) > 0 Then
       sSQL = sSQL & _
              "AND  (SUBSTRING(pallet, 5, 4) + SUBSTRING(pallet, 3, 2) + SUBSTRING(pallet, 1, 2)) >= '" & cFields(1) & "' "
       sSQL = sSQL & _
              "AND SUBSTRING(pallet, 10, 3) >= '" & cFields(3) & "' "
    End If
    
    If Len(Trim(cFields(2))) > 0 Then
       sSQL = sSQL & _
              "AND  (SUBSTRING(pallet, 5, 4) + SUBSTRING(pallet, 3, 2) + SUBSTRING(pallet, 1, 2)) <= '" & cFields(2) & "' "
       sSQL = sSQL & _
              "AND SUBSTRING(pallet, 10, 3) <= '" & cFields(4) & "' "
    End If

    If cFields(5) = "1" Then 'FATURADO
       sSQL = sSQL & _
              "AND  (CONF_PALLET = 'P' OR CONF_PALLET = 'F') " & _
              " "
    ElseIf cFields(5) = "2" Then 'EMBARCOU
       sSQL = sSQL & _
              "AND  CONF_PALLET = 'E' "
    ElseIf cFields(5) = "3" Then 'NAO FATURADO(NAO PALETIZADO)
       sSQL = sSQL & _
              "AND  CONF_PALLET = '' " & _
              "AND  XBLNR > 0 "
    ElseIf cFields(5) = "4" Then 'NAO EMBARCOU (NA FABRICA)
       sSQL = sSQL & _
              " AND  CONF_PALLET <> 'E' AND XBLNR > 0 "
    End If

    If Len(Trim(cFields(6))) > 0 Then
       sSQL = sSQL & _
              " AND  ID_CLIENTE = '" & cFields(6) & "' "
    End If

    If Len(Trim(cFields(7))) > 0 Then
       sSQL = sSQL & _
              " AND  XBLNR >= " & cFields(7)
    End If
    
    If Len(Trim(cFields(8))) > 0 Then
       sSQL = sSQL & _
              " AND  XBLNR <= " & cFields(8)
    End If
    
    sSQL = sSQL & _
           " GROUP BY PALLET,ID_CLIENTE,CONF_PALLET,FATUR "
    
'****************************************************************************************************
'********** Consultar ETIQUETAS PELO NUMERO DO CLIENTE/PALLET/INTERVALO DE NF
'****************************************************************************************************
    Case 13
    
    sSQL = "SELECT " & _
           "ID_ETIQUETA AS NUM_CAIXA, " & _
           "CASE TIPO_PADRAO_CX WHEN 1 THEN 'IK10' ELSE " & _
           "CASE TIPO_PADRAO_CX WHEN 2 THEN 'IK33' ELSE 'PAPELAO' " & _
           "END END AS TIPO_CAIXA, " & _
           "ID_PECA AS PECA," & _
           "LOTE," & _
           "QTDE AS QTDE_NA_CAIXA," & _
           "XBLNR AS NF_VENDA, " & _
           "REMESSA AS ORDEM_VENDA," & _
           "SEQUENCIA_PLACA AS SEQUENCIA, " & _
           "PLACA "
    sSQL = sSQL & _
           "FROM ETIQUETA "
           
    sSQL = sSQL & _
           "WHERE  ID_CLIENTE = '" & cFields(4) & "' "
    
    sSQL = sSQL & _
           "AND  PALLET = '" & cFields(1) & "' "
              
    If Len(Trim(cFields(2))) > 0 Then
       sSQL = sSQL & _
              "AND XBLNR >= " & cFields(2)
    End If

    If Len(Trim(cFields(3))) > 0 Then
       sSQL = sSQL & _
              "AND XBLNR <= " & cFields(3)
    End If
    
    If cFields(5) = "1" Then 'FATURADO
       sSQL = sSQL & _
              " AND (CONF_PALLET = 'P' OR CONF_PALLET = 'F') AND " & _
              "FATUR <> '' ": bWhere = True
    ElseIf cFields(5) = "2" Then 'EMBARCOU
       sSQL = sSQL & _
              " AND CONF_PALLET = 'E' ": bWhere = True
    ElseIf cFields(5) = "3" Then 'NAO FATURADO
       sSQL = sSQL & _
              " AND (CONF_PALLET = 'P' OR CONF_PALLET = '') AND " & _
              "FATUR = '' ": bWhere = True
    ElseIf cFields(5) = "4" Then 'NAO EMBARCOU (NA FABRICA)
       sSQL = sSQL & _
              " AND (CONF_PALLET = 'P' OR CONF_PALLET = '') ": bWhere = True
    End If
    
    sSQL = sSQL & _
           " ORDER BY ID_PECA "
    
'****************************************************************************************************
'********** consultar as pecas existentes na tabela de etiquets
'****************************************************************************************************
    Case 14 'Consultar pecas na tab etiqueta
    
    sSQL = "SELECT " & _
           "DISTINCT ID_PECA "
    sSQL = sSQL & _
           "FROM ETIQUETA "
    If Len(Trim(cFields(1))) > 0 Then
       sSQL = sSQL & _
              "WHERE  ID_CLIENTE = '" & cFields(1) & "' "
    End If
    
    sSQL = sSQL & _
           " ORDER BY ID_PECA "

'****************************************************************************************************
'********** consultar as caixa existente na tabela de etiquetas
'****************************************************************************************************
    Case 15 'Consultar caixa na tab etiqueta
    
    sSQL = "SELECT " & _
           "ID_ETIQUETA "
    sSQL = sSQL & _
           "FROM ETIQUETA "
    sSQL = sSQL & _
           "WHERE ID_ETIQUETA = '" & cFields(1) & "' "
    If Len(Trim(cFields(2))) > 0 Then
       sSQL = sSQL & _
              " AND ID_CLIENTE = '" & cFields(2) & "' "
    End If

End Select

Set rs = New ADODB.Recordset
If Len(Trim(sSQL)) > 0 Then
   cCONECCAO.CursorLocation = adUseClientBatch
   rs.Open sSQL, cCONECCAO
End If

Set EXPEDICAO_dados = rs

Set rs = Nothing
Set rsSeq = Nothing

Exit Function

Erro:
    Set rs = Nothing
    Set rsSeq = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function




