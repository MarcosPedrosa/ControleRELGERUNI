VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Variável pública que conterá o MDIform
'Option Explicit
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS DO TURMA
' ***********************************************************
' ***********************************************************
' ***********************************************************

Public Function USUARIO_Incluir(ByVal sBanco As String, _
                                ByVal c_USU_CODIGO As String, _
                                ByVal c_USU_USUARIO As String, _
                                ByVal c_USU_SENHA As String, _
                                ByVal c_USUARIO As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_USU_CODIGO, "c_USU_CODIGO"
cFields.Add c_USU_USUARIO, "c_USU_USUARIO"
cFields.Add c_USU_SENHA, "c_USU_SENHA"
cFields.Add c_USUARIO, "c_USUARIO"

Set USUARIO_Incluir = USUARIO_Critica(ADOConnection, cFields, 2)

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

Public Function USUARIO_Consultar_Acesso(ByVal sBanco As String, _
                                Optional ByVal c_USU_CODIGO As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

If IsNull(c_USU_CODIGO) Then
Else
   cFields.Add c_USU_CODIGO, "c_USU_CODIGO"
End If

Set USUARIO_Consultar_Acesso = USUARIO_Critica(ADOConnection, cFields, 11)

If USUARIO_Consultar_Acesso.RecordCount = 0 Then
   Set USUARIO_Consultar_Acesso = USUARIO_Critica(ADOConnection, cFields, 12)
   Set USUARIO_Consultar_Acesso = USUARIO_Critica(ADOConnection, cFields, 11)
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
Public Function USUARIO_Permissao_Alterar(ByVal sBanco As String, _
                                          ByVal c_FRM_FOR_CODIGO As String, _
                                          ByVal c_FRM_VISUALIZA As String, _
                                          ByVal c_FRM_HABILITAR As String, _
                                          ByVal c_USUARIO As String, _
                                          ByVal c_DATA_ALT As String, _
                                          ByVal c_FRM_GRUPO As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_FRM_FOR_CODIGO, "c_FRM_FOR_CODIGO" '1
cFields.Add c_FRM_VISUALIZA, "c_FRM_VISUALIZA" '2
cFields.Add c_FRM_HABILITAR, "c_FRM_HABILITAR" '3
cFields.Add c_USUARIO, "c_USUARIO" '4
cFields.Add c_DATA_ALT, "c_DATA_ALT" '5
cFields.Add c_FRM_GRUPO, "c_FRM_GRUPO" '6

Set USUARIO_Permissao_Alterar = USUARIO_Critica(ADOConnection, cFields, 13)

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

Public Function USUARIO_Consultar(ByVal sBanco As String, _
                                  Optional ByVal c_USU_CODIGO As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

If IsNull(c_USU_CODIGO) Then
Else
   cFields.Add c_USU_CODIGO, "c_USU_CODIGO"
End If

Set USUARIO_Consultar = USUARIO_Critica(ADOConnection, cFields, 1)

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
Public Function USUARIO_Alterar(ByVal sBanco As String, _
                                ByVal c_USU_CODIGO As String, _
                                ByVal c_USU_USUARIO As String, _
                                ByVal c_USU_SENHA As String, _
                                ByVal c_USUARIO As String, _
                                ByVal c_DATA_ALT As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_USU_CODIGO, "c_USU_CODIGO" '1
cFields.Add c_USU_USUARIO, "c_USU_USUARIO" '2
cFields.Add c_USU_SENHA, "c_USU_SENHA" '18
cFields.Add c_USUARIO, "c_USUARIO" '18
cFields.Add c_DATA_ALT, "c_DATA_ALT" '19

Set USUARIO_Alterar = USUARIO_Critica(ADOConnection, cFields, 3)

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
Public Function USUARIO_Excluir(ByVal sBanco As String, _
                                ByVal c_USU_CODIGO As Integer, _
                              ByVal c_DATA_ALT As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_USU_CODIGO, "c_USU_CODIGO" '1
cFields.Add Null, "c_USU_USUARIO" '2
cFields.Add c_USU_SENHA, "c_USU_SENHA" '18
cFields.Add Null, "c_USUARIO" '18
cFields.Add c_DATA_ALT, "c_DATA_ALT" '19

Set USUARIO_Excluir = USUARIO_Critica(ADOConnection, cFields, 4)

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
Public Function USUARIO_Inserir_Acesso(ByVal cCONECCAO As Connection, _
                                       ByVal cFields As Collection) As ADODB.Recordset

On Error GoTo Erro

Set USUARIO_Inserir_Acesso = USUARIO_Critica(cCONECCAO, cFields, 14)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function USUARIO_Alterar_Acesso(ByVal cCONECCAO As Connection, _
                                       ByVal cFields As Collection) As ADODB.Recordset

On Error GoTo Erro

Set USUARIO_Alterar_Acesso = USUARIO_Critica(cCONECCAO, cFields, 15)

Set USUARIO_Alterar_Acesso = USUARIO_Critica(cCONECCAO, cFields, 14)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function USUARIO_Excluir_Acesso(ByVal cCONECCAO As Connection, _
                                       ByVal cFields As Collection) As ADODB.Recordset

On Error GoTo Erro

Set USUARIO_Excluir_Acesso = USUARIO_Critica(cCONECCAO, cFields, 15)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function USUARIO_Critica(ByVal cCONECCAO As Connection, _
                                ByVal cFields As Collection, _
                                ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction > 1 And nAction < 11 Then
   'Critica dos Campos dos TURMA
    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
       Err.Raise 50000, Err.Source, "Código do TURMA Inválido"
    End If
    
    If Len(Trim(cFields(2))) = 0 Then
       Err.Raise 50000, Err.Source, "Nome do TURMA Inválido"
    End If
   
    If Len(Trim(cFields(3))) = 0 Then
       Err.Raise 50000, Err.Source, "Codigo do PROFESSOR Inválido"
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

Set USUARIO_Critica = USUARIO_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE DADOS DO TURMA
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function USUARIO_dados(ByVal cCONECCAO As Connection, _
                              ByVal cFields As Collection, _
                              ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String
Dim nseq As Integer
Dim rsSeq As ADODB.Recordset


Select Case nAction
    Case 1 'Consultar TURMA
    
    sSQL = "SELECT   USU_CODIGO, " & _
                    "USU_USUARIO, " & _
                    "USU_SENHA, " & _
                    "USU_USU, " & _
                    "USU_DTA, " & _
                    "USU_DTI " & _
                    "FROM EXP_TAB_USUARIO "
    
    If cFields.Count > 0 Then
       If Len(Trim(cFields(1))) > 0 Then
          sSQL = sSQL & " WHERE USU_CODIGO = " & cFields(1)
       Else
          sSQL = sSQL & " ORDER BY USU_USUARIO"
       End If
    Else
       sSQL = sSQL & " ORDER BY USU_USUARIO"
    End If
    
    Case 2 'Incluir TURMA
    
    sDate = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"

    sSQL = "INSERT INTO EXP_TAB_USUARIO (" & _
                    "" & _
                    "USU_USUARIO, " & _
                    "USU_SENHA, " & _
                    "USU_USU, " & _
                    "USU_DTA, " & _
                    "USU_DTI "
                    
    sSQL = sSQL & ") VALUES (" & _
                  cFields(2) & ", " & _
                  cFields(3) & ", " & cFields(4) & ", " & _
                  sDate & ", " & sDate & ")"
                    
    Case 3 'Alterar TURMA
    
    sDate = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"
'                 "  EXP_TAB_USUARIO.USU_CODIGO = " & cFields(1)
    sSQL = "UPDATE EXP_TAB_USUARIO SET" & _
                 "  EXP_TAB_USUARIO.USU_USUARIO = " & cFields(2) & _
                 ", EXP_TAB_USUARIO.USU_SENHA = " & cFields(3) & _
                 ", EXP_TAB_USUARIO.USU_USU = " & cFields(4) & _
                 ", EXP_TAB_USUARIO.USU_DTA = " & sDate & _
                 "  WHERE EXP_TAB_USUARIO.USU_CODIGO = " & cFields(1)
    
    Case 4 'Excluir TURMA
    sSQL = "DELETE * FROM EXP_TAB_USUARIO" & _
                 "  WHERE EXP_TAB_USUARIO.USU_CODIGO = " & cFields(1) & _
                 " AND EXP_TAB_USUARIO.USU_DTA = " & cFields(5)
                 
                 
    Case 11 'Consultar Acesso dos usuarios
    
    sSQL = "SELECT " & _
           "EXP_REL_USUARIO_FORMULARIO.FRM_VISUALIZA,  " & _
           "EXP_REL_USUARIO_FORMULARIO.FRM_HABILITAR,  " & _
           "EXP_TAB_FORMULARIO.FOR_NOME_EDITOR,  " & _
           "EXP_TAB_FORMULARIO.FOR_NOME_FORM,  " & _
           "EXP_TAB_FORMULARIO.FOR_DESCRICAO,  " & _
           "EXP_TAB_FORMULARIO.FOR_GRUPO,  " & _
           "EXP_REL_USUARIO_FORMULARIO.FRM_BOTAO,  " & _
           "EXP_REL_USUARIO_FORMULARIO.FRM_ACESSO,  " & _
           "EXP_REL_USUARIO_FORMULARIO.FRM_DTA,  " & _
           "EXP_REL_USUARIO_FORMULARIO.FRM_DTI,  " & _
           "EXP_REL_USUARIO_FORMULARIO.FRM_USU, " & _
           "EXP_REL_USUARIO_FORMULARIO.FRM_FOR_CODIGO "
    sSQL = sSQL & _
           "FROM EXP_REL_USUARIO_FORMULARIO  " & _
           "INNER JOIN EXP_TAB_FORMULARIO  " & _
           "ON EXP_REL_USUARIO_FORMULARIO.FRM_FOR_CODIGO = EXP_TAB_FORMULARIO.FOR_CODIGO "

    If cFields.Count > 0 Then
       If Len(Trim(cFields(1))) > 0 Then
          sSQL = sSQL & " WHERE FRM_USU_CODIGO = " & cFields(1)
          sSQL = sSQL & " ORDER BY FOR_GRUPO"
       Else
          sSQL = sSQL & " ORDER BY FOR_GRUPO"
       End If
    Else
       sSQL = sSQL & " ORDER BY FOR_GRUPO"
    End If

    Case 12 'Incluir Acessos para usuario
    
    sDate = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"

    sSQL = "INSERT INTO EXP_REL_USUARIO_FORMULARIO ( " & _
           "FRM_USU_CODIGO,  " & _
           "FRM_FOR_CODIGO,  " & _
           "FRM_VISUALIZA,  " & _
           "FRM_HABILITAR,  " & _
           "FRM_BOTAO,  " & _
           "FRM_ACESSO,  " & _
           "FRM_DTA,  " & _
           "FRM_DTI,  " & _
           "FRM_USU ) "
           
    sSQL = sSQL & _
           "SELECT " & cFields(1) & " AS Expr1,  " & _
           "EXP_TAB_FORMULARIO.FOR_CODIGO,  " & _
           "0 ,  " & _
           "0 ,  " & _
           "1111 ,  " & _
           "EXP_TAB_FORMULARIO.FOR_ACESSO,  " & _
           sDate & " ,  " & _
           sDate & " ,  " & _
           "100 " & _
           "FROM EXP_TAB_FORMULARIO"

    Case 13 'Alterar Acessos para usuario
    
    sSQL = "SELECT FRM_FOR_CODIGO FROM EXP_REL_USUARIO_FORMULARIO " & _
                 "  WHERE EXP_REL_USUARIO_FORMULARIO.FRM_FOR_CODIGO = " & cFields(1)

    Set rs = New ADODB.Recordset
    cCONECCAO.CursorLocation = adUseClientBatch
    rs.Open sSQL, cCONECCAO
    
    If rs.RecordCount = 0 Then
       Err.Raise 50000, Err.Source, "Erro na atualização do acesso do usuário,Tente novamente!"
    End If
    sDate = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"

    sSQL = "UPDATE EXP_REL_USUARIO_FORMULARIO SET" & _
                 "  FRM_VISUALIZA = " & cFields(2) & _
                 ", FRM_HABILITAR = " & cFields(3) & _
                 ", FRM_DTA = " & sDate & _
                 ", FRM_USU = " & cFields(4) & _
                 "  WHERE EXP_REL_USUARIO_FORMULARIO.FRM_FOR_CODIGO = " & cFields(1) & _
                 "  AND   EXP_REL_USUARIO_FORMULARIO.FRM_USU_CODIGO = " & cFields(4)

    Case 14 'Incluir o Formulario novo com os Acessos para o usuario
    
    sDate = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"

    sSQL = "INSERT INTO EXP_REL_USUARIO_FORMULARIO ( " & _
           "FRM_USU_CODIGO,  " & _
           "FRM_FOR_CODIGO,  " & _
           "FRM_VISUALIZA,  " & _
           "FRM_HABILITAR,  " & _
           "FRM_BOTAO,  " & _
           "FRM_ACESSO,  " & _
           "FRM_DTA,  " & _
           "FRM_DTI,  " & _
           "FRM_USU ) "
           
    sSQL = sSQL & _
           "SELECT USU_CODIGO ,  " & _
           cFields(1) & ",  " & _
           "0 ,  " & _
           "0 ,  " & _
           "'1111', '" & _
           cFields(6) & "',  " & _
           sDate & " ,  " & _
           sDate & " ,  " & _
           cFields(7) & " " & _
           "FROM EXP_TAB_USUARIO"
           
    Case 15 'Excluir o Formulario com os Acessos para o usuario
    
    sDate = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"

    sSQL = "DELETE  FROM EXP_REL_USUARIO_FORMULARIO  " & _
           "WHERE FRM_FOR_CODIGO =  " & cFields(1)

End Select

Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch

rs.Open sSQL, cCONECCAO

If nAction = 2 Then  'Incluir TURMA
   sSQL = "SELECT MAX(EXP_TAB_USUARIO.USU_CODIGO) AS ultimoDeCLI_CODIGO" & _
                 " FROM EXP_TAB_USUARIO;"
   cCONECCAO.CursorLocation = adUseClientBatch
   rs.Open sSQL, cCONECCAO
End If

Set USUARIO_dados = rs

Set rs = Nothing
Set rsSeq = Nothing

Exit Function

Erro:
    Set rs = Nothing
    Set rsSeq = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function






