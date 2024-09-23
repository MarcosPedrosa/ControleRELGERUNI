VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neManutencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS DO EXPEDICAO
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public sObs_Log As String  'contera a observacao para gravar o log

Public Function MANUTENCAO_NF_Consultar(ByVal sBanco As String, _
                                        ByVal sNF As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sNF, "sNF"

Set MANUTENCAO_NF_Consultar = MANUTENCAO_Critica(ADOConnection, cFields, 1)

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
Public Function MANUTENCAO_NF_Cancelar(ByVal sBanco As String, _
                                       ByVal sNF As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sNF, "sNF"

Set MANUTENCAO_NF_Cancelar = MANUTENCAO_Critica(ADOConnection, cFields, 2)

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
Public Function MANUTENCAO_CAIXA_Consultar(ByVal sBanco As String, _
                                           ByVal sCX As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sCX, "sCX"

Set MANUTENCAO_CAIXA_Consultar = MANUTENCAO_Critica(ADOConnection, cFields, 3)

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
Public Function MANUTENCAO_CAIXA_Alterar_Tipo(ByVal sBanco As String, _
                                              ByVal sCX As String, _
                                              ByVal sTipo As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sCX, "sCX"
cFields.Add sTipo, "sTipo"

Set MANUTENCAO_CAIXA_Alterar_Tipo = MANUTENCAO_Critica(ADOConnection, cFields, 4)

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
Public Function MANUTENCAO_INVENTARIO_Qt_Registro(ByVal sBanco As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set MANUTENCAO_INVENTARIO_Qt_Registro = MANUTENCAO_Critica(ADOConnection, cFields, 5)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set cFields = Nothing
Set ADOConnection = Nothing
Set cConect = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function MANUTENCAO_INVENTARIO_Abertura(ByVal sBanco As String, _
                                               ByVal sUsuario As String, _
                                               ByVal sdata As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sUsuario, "sUsuario"
cFields.Add sdata, "sData"
cFields.Add "ETIQUETA", "1"
cFields.Add "ABERTURA INVENTARIO", "2"

Rem informacoes de observacoes para gravacao do registro de log
sObs_Log = "Rotina de Abertura de inventário, no dia " & sdata & " pelo usuário " & Trim(sUsuario) & " às " & Format(Now(), "HH:MM:SS") & "."

Set MANUTENCAO_INVENTARIO_Abertura = MANUTENCAO_Critica(ADOConnection, cFields, 6)

sObs_Log = ""

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
Public Function MANUTENCAO_Consultar_Mov_Para_R3(ByVal sBanco As String, _
                                                ByVal sDataINI As String, _
                                                ByVal sDataFIM As String, _
                                                ByVal sSEQ1 As String, _
                                                ByVal sSEQ2 As String) As ADODB.Recordset

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

Set MANUTENCAO_Consultar_Mov_Para_R3 = MANUTENCAO_Critica(ADOConnection, cFields, 7)

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
Public Function MANUTENCAO_INVENTARIO_Importar(ByVal sBanco As String, _
                                               ByVal sUsuario As String, _
                                               ByVal sdata As String, _
                                               ByVal sArquivo As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Rem deletar arquivo temporario
Rem adicionar registros vindo do R3

cFields.Add sUsuario, "sUsuario"
cFields.Add sdata, "sData"
cFields.Add "ETIQUETA", "1"
cFields.Add "IMPORTACAO INVENTARIO", "2"
cFields.Add sArquivo, "sArquivo"

Call MANUTENCAO_Critica(ADOConnection, cFields, 8)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
'Close #11


Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Close #11
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function MANUTENCAO_INVENTARIO_GeraPaletR3(ByVal sBanco As String, _
                                                  ByVal sUsuario As String, _
                                                  ByVal sdata As String, _
                                                  ByVal sLido As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sUsuario, "sUsuario"
cFields.Add sdata, "sData"
cFields.Add "ETIQUETA", "1"
cFields.Add "GERAR PALLET R3,INVENTARIO", "2"

Rem informacoes de observacoes para gravacao do registro de log
sObs_Log = "Rotina de Gerar Pallets para R3, no dia " & sdata & " pelo usuário " & Trim(sUsuario) & " às " & Format(Now(), "HH:MM:SS") & ". Lidos " & sLido & " registros."

Call MANUTENCAO_Critica(ADOConnection, cFields, 0)

sObs_Log = ""

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

Public Function MANUTENCAO_Critica(ByVal cCONECCAO As Connection, _
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

Set MANUTENCAO_Critica = MANUTENCAO_dados(cCONECCAO, cFields, nAction)

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
Public Function MANUTENCAO_dados(ByVal cCONECCAO As Connection, _
                                 ByVal cFields As Collection, _
                                 ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String
Dim bWhere As Boolean
Dim sNome_arquivo As String
Dim Y As Double
Dim x As Double
Dim nseq As Double

Select Case nAction
'****************************************************************************************************
'********** consultar nota fiscal para posterior cancelamento
'****************************************************************************************************
    Case 1
    
    sSQL = "SELECT " & _
           "ID_ETIQUETA, " & _
           "ID_PECA, " & _
           "QTDE, " & _
           "LOTE, " & _
           "CASE TIPO_PADRAO_CX WHEN 1 THEN 'IK10' ELSE " & _
           "CASE TIPO_PADRAO_CX WHEN 2 THEN 'IK33' ELSE 'PAPELAO' " & _
           "END END AS TIPO_CAIXA, " & _
           "PALLET," & _
           "PLACA "
           
    sSQL = sSQL & _
           "FROM ETIQUETA "
           
    sSQL = sSQL & _
           "WHERE  PALLET = '' "
    
    If Len(Trim(cFields(1))) > 0 Then
       sSQL = sSQL & _
              "AND  XBLNR = '" & cFields(1) & "' "
    End If

'****************************************************************************************************
'********** Cancelar nota fiscal com seus itens
'****************************************************************************************************
    Case 2
    
    sSQL = "UPDATE ETIQUETA SET " & _
           "REMESSA = '', " & _
           "FATUR = '', " & _
           "XBLNR = '', " & _
           "POSNR = 0 "
           
    sSQL = sSQL & _
           "WHERE  PALLET = '' "
    
    If Len(Trim(cFields(1))) > 0 Then
       sSQL = sSQL & _
              "AND  XBLNR = '" & cFields(1) & "' "
    End If
    
'****************************************************************************************************
'********** consultar caixa para posterior mudanca de tipo da caixa
'****************************************************************************************************
    Case 3
    
    sSQL = "SELECT " & _
           "ID_ETIQUETA, " & _
           "CASE TIPO_PADRAO_CX WHEN 1 THEN 'IK10' ELSE " & _
           "CASE TIPO_PADRAO_CX WHEN 2 THEN 'IK33' ELSE 'PAPELAO' " & _
           "END END AS TIPO_CAIXA "
           
    sSQL = sSQL & _
           "FROM ETIQUETA "
    
    If Len(Trim(cFields(1))) > 0 Then
       sSQL = sSQL & _
              "WHERE ID_ETIQUETA = " & cFields(1)
    End If

'****************************************************************************************************
'********** Alterar o tipo da caixa
'****************************************************************************************************
    Case 4
    
    sSQL = "UPDATE " & _
           "ETIQUETA SET " & _
           "TIPO_PADRAO_CX = " & cFields(2)
    
    sSQL = sSQL & _
           " WHERE ID_ETIQUETA = " & cFields(1)
           
'****************************************************************************************************
'********** Quantidade de registros que a abertura de inventario vai atualizar
'****************************************************************************************************
    Case 5
    
    sSQL = "SELECT " & _
           "COUNT(*) AS TOTAL_REG," & _
           "(SELECT COUNT(*) " & _
           "FROM ETIQUETA WHERE ID_BORDERO <> '') AS INVENTARIADO FROM ETIQUETA "
           
'****************************************************************************************************
'********** Abertura de inventario
'****************************************************************************************************
    Case 6
    
    sSQL = "UPDATE ETIQUETA SET " & _
           "ID_BORDERO = '' " & _
           "WHERE ID_BORDERO <> ''"
           
'****************************************************************************************************
'********** Consultar OS PALLETS PARA INFORMAR AO R3
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
           "PALLET "
    sSQL = sSQL & _
           "FROM ETIQUETA "
    sSQL = sSQL & _
           "WHERE  ID_BORDERO <> '' "
    
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
              
    sSQL = sSQL & _
           " ORDER BY ID_ETIQUETA "
           
'****************************************************************************************************
'********** importacao de mov. inventario (DELECAO Arquivo temporario)/inclusao de registros vindo do R3
'****************************************************************************************************
    Case 8
    sSQL = "DELETE FROM EXP_TMP_INVENTARIO "

    Set rs = New ADODB.Recordset
           
    cCONECCAO.CursorLocation = adUseClientBatch
    rs.Open sSQL, cCONECCAO
    
'***********************************************************************************************

    sNome_arquivo = cFields(5)
    
    Close #11
    Open sNome_arquivo For Random Access Read Write As #11 Len = Len(Arq_Mov_Inventario)
    
    Y = LOF(11) / Len(Arq_Mov_Inventario)
    
    frmGIFMntImportaMovInventário.Refresh
    frmGIFMntImportaMovInventário.txtlidos.Text = Y
    
    nAtualizadas = 0
    frmGIFMntImportaMovInventário.PBar1.Value = nseq
    frmGIFMntImportaMovInventário.PBar1.Visible = True
    frmGIFMntImportaMovInventário.PBar1.Min = 0
    Y = LOF(11) / Len(Arq_Mov_Inventario)
    frmGIFMntImportaMovInventário.PBar1.Max = Y
    x = 0
    nseq = 0

    For Y = 1 To LOF(11) / Len(Arq_Mov_Inventario)
        nseq = nseq + 1
        frmGIFMntImportaMovInventário.PBar1.Value = nseq
        Get 11, Y, Arq_Mov_Inventario
        nAtualizadas = nAtualizadas + 1
        sSQL = "INSERT INTO EXP_TMP_INVENTARIO (ID_ETIQUETA,ID_BORDERO) VALUES ('" & _
               Trim(Arq_Mov_Inventario.FRegistro) & "','" & _
               Trim(Arq_Mov_Inventario.FTipo) & "')"
        cCONECCAO.CursorLocation = adUseClientBatch
        rs.Open sSQL, cCONECCAO
        frmGIFMntImportaMovInventário.txt_atualizado.Text = nseq
        frmGIFMntImportaMovInventário.txt_atualizadas.Text = nAtualizadas
        frmGIFMntImportaMovInventário.txt_atualizado.Refresh
        frmGIFMntImportaMovInventário.txt_atualizadas.Refresh
    Next

'***********************************************************************************************
    sSQL = "SELECT COUNT(*) AS CONTA" & _
           "FROM EXP_TMP_INVENTARIO INV " & _
           "INNER JOIN ETIQUETA ETI " & _
           "ON INV.ID_ETIQUETA = ETI.ID_ETIQUETA AND INV.ID_BORDERO <> ETI.ID_BORDERO "
        cCONECCAO.CursorLocation = adUseClientBatch
        rs.Open sSQL, cCONECCAO
    
    sObs_Log = "Rotina de Importação de etiqueta do banco R3 para SQL, no dia " & cFields(5) & " pelo usuário " & Trim(cFields(5)) & " às " & Format(Now(), "HH:MM:SS") & ". Lendo - " & Format(Y, "0") & " Registros, Gravando - " & Format(rs!CONTA, "0") & " Registros."
    
    sSQL = "UPDATE ETI " & _
           "SET ETI.ID_BORDERO = INV.ID_BORDERO " & _
           "FROM EXP_TMP_INVENTARIO INV " & _
           "INNER JOIN ETIQUETA ETI " & _
           "ON INV.ID_ETIQUETA = ETI.ID_ETIQUETA AND INV.ID_BORDERO <> ETI.ID_BORDERO "
    
End Select

Set rs = New ADODB.Recordset

If Len(Trim(sSQL)) > 0 Then
   cCONECCAO.CursorLocation = adUseClientBatch
   rs.Open sSQL, cCONECCAO
End If

Set MANUTENCAO_dados = rs

'Gravar registro de Log
If sObs_Log <> Empty Then
   Call CCTempneLog.LOG_INCLUSAO(cCONECCAO, _
                                 cFields(2), _
                                 Format(Now(), "HH:MM:SS"), _
                                 cFields(1), _
                                 cFields(3), _
                                 cFields(4), _
                                 Replace(sSQL, "'", "::"), _
                                 sObs_Log)
End If
'Gravar registro de Log (FIM)

Set rs = Nothing

Exit Function

Erro:
    Set rs = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

