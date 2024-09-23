VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neFormulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Variável pública que conterá o MDIform
'Option Explicit
'**************** Ações Tipos de Formularios
'Consultar Tipos de Formularios = 1
'Incluir Tipos de Formularios = 2
'Alterar Tipos de Formularios = 3
'Excluir Tipos de Formularios = 4
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS DO Feriado
' ***********************************************************
' ***********************************************************
' ***********************************************************

Public Function TipoForm_Incluir(ByVal sBanco As String, _
                                 ByVal c_FOR_CODIGO As String, _
                                 ByVal c_FOR_NOME_EDITOR As String, _
                                 ByVal c_FOR_NOME_FORM As String, _
                                 ByVal c_FOR_DESCRICAO As String, _
                                 ByVal c_FOR_GRUPO As String, _
                                 ByVal c_FOR_ACESSO As String, _
                                 ByVal c_USUARIO As String) As ADODB.Recordset

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_FOR_CODIGO, "c_FOR_CODIGO"
cFields.Add c_FOR_NOME_EDITOR, "c_FOR_NOME_EDITOR"
cFields.Add c_FOR_NOME_FORM, "c_FOR_NOME_FORM"
cFields.Add c_FOR_DESCRICAO, "c_FOR_DESCRICAO"
cFields.Add c_FOR_GRUPO, "c_FOR_GRUPO"
cFields.Add c_FOR_ACESSO, "c_FOR_ACESSO"
cFields.Add c_USUARIO, "c_USUARIO"

Set TipoForm_Incluir = TipoForm_Critica(ADOConnection, cFields, 2)

Set cFields = New Collection

cFields.Add TipoForm_Incluir.Fields(0), "c_FOR_CODIGO"
cFields.Add c_FOR_NOME_EDITOR, "c_FOR_NOME_EDITOR"
cFields.Add c_FOR_NOME_FORM, "c_FOR_NOME_FORM"
cFields.Add c_FOR_DESCRICAO, "c_FOR_DESCRICAO"
cFields.Add c_FOR_GRUPO, "c_FOR_GRUPO"
cFields.Add c_FOR_ACESSO, "c_FOR_ACESSO"
cFields.Add c_USUARIO, "c_USUARIO"

Call CCTempneUsuario.USUARIO_Inserir_Acesso(ADOConnection, cFields)

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

Public Function TipoForm_Consultar(ByVal sBanco As String, _
                          Optional ByVal c_FOR_CODIGO As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

If Len(Trim(c_FOR_CODIGO)) > 0 Then
   cFields.Add c_FOR_CODIGO, "c_FOR_CODIGO"
End If

Set TipoForm_Consultar = TipoForm_Critica(ADOConnection, cFields, 1)

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
Public Function TipoForm_Alterar(ByVal sBanco As String, _
                                 ByVal c_FOR_CODIGO As String, _
                                 ByVal c_FOR_NOME_EDITOR As String, _
                                 ByVal c_FOR_NOME_FORM As String, _
                                 ByVal c_FOR_DESCRICAO As String, _
                                 ByVal c_FOR_GRUPO As String, _
                                 ByVal c_FOR_ACESSO As String, _
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

Set cFields = New Collection
cFields.Add c_FOR_CODIGO, "c_FOR_CODIGO" '1
cFields.Add c_FOR_NOME_EDITOR, "c_FOR_NOME_EDITOR" '2
cFields.Add c_FOR_NOME_FORM, "c_FOR_NOME_FORM" '3
cFields.Add c_FOR_DESCRICAO, "c_FOR_DESCRICAO" '4
cFields.Add c_FOR_GRUPO, "c_FOR_GRUPO" '5
cFields.Add c_FOR_ACESSO, "c_FOR_ACESSO" '6
cFields.Add c_USUARIO, "c_USUARIO" '7
cFields.Add c_DATA_ALT, "c_DATA_ALT" '8

Set TipoForm_Alterar = TipoForm_Critica(ADOConnection, cFields, 3)

Call CCTempneUsuario.USUARIO_Alterar_Acesso(ADOConnection, cFields)

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
Public Function TipoForm_Excluir(ByVal sBanco As String, _
                                 ByVal c_FOR_CODIGO As String, _
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

Set cFields = New Collection
cFields.Add c_FOR_CODIGO, "c_FOR_CODIGO" '1
cFields.Add Null, "c_FOR_NOME_EDITOR" '2
cFields.Add Null, "c_FOR_NOME_FORM" '3
cFields.Add Null, "c_FOR_DESCRICAO" '4
cFields.Add Null, "c_FOR_GRUPO" '5
cFields.Add Null, "c_FOR_ACESSO" '6
cFields.Add c_USUARIO, "c_USUARIO" '7
cFields.Add c_DATA_ALT, "c_DATA_ALT" '8

Set TipoForm_Excluir = TipoForm_Critica(ADOConnection, cFields, 4)

Call CCTempneUsuario.USUARIO_Excluir_Acesso(ADOConnection, cFields)

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
Public Function TipoForm_Critica(ByVal cCONECCAO As Connection, _
                                 ByVal cFields As Collection, _
                                 ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction > 1 Then
   'Critica dos Campos dos Tipos de Formularios
    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
       Err.Raise 50000, Err.Source, "Código do Posto Inválido"
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
            Case 2, 3, 4, 5, 6, 7
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
                Case 2, 3, 4, 5
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
                Case 8
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
                Case 0
'                    If InStr(cFields(nx), ",") > 0 Then
'                        sSQL = Replace(cFields(NX), ".", "")
'                        sSQL = Replace(sSQL, ",", ".")
'                        cFields.Remove nx
'                        cFields.Add sSQL, , , nx - 1
'                    End If
            End Select
        End If
    
    Next nx
End If

Set TipoForm_Critica = TipoForm_dados(cCONECCAO, cFields, nAction)

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
Public Function TipoForm_dados(ByVal cCONECCAO As Connection, _
                               ByVal cFields As Collection, _
                               ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String

Select Case nAction
    Case 1 'Consultar Tipos de Formularios

    sSQL = "SELECT " & _
           "EXP_TAB_FORMULARIO.FOR_CODIGO, " & _
           "EXP_TAB_FORMULARIO.FOR_NOME_EDITOR, " & _
           "EXP_TAB_FORMULARIO.FOR_NOME_FORM, " & _
           "EXP_TAB_FORMULARIO.FOR_DESCRICAO, " & _
           "EXP_TAB_FORMULARIO.FOR_GRUPO, " & _
           "EXP_TAB_FORMULARIO.FOR_ACESSO, " & _
           "EXP_TAB_FORMULARIO.FOR_DTA, " & _
           "EXP_TAB_FORMULARIO.FOR_DTI, " & _
           "EXP_TAB_FORMULARIO.FOR_USU "
           
    sSQL = sSQL & _
           " FROM EXP_TAB_FORMULARIO"

            
    If cFields.Count > 0 Then
       If Len(Trim(cFields(1))) > 0 Then
          sSQL = sSQL & " WHERE FOR_CODIGO = " & cFields(1) & ""
       Else
          sSQL = sSQL & " ORDER BY FOR_GRUPO,FOR_NOME_EDITOR,FOR_ACESSO"
       End If
    Else
       sSQL = sSQL & " ORDER BY FOR_GRUPO,FOR_NOME_EDITOR,FOR_ACESSO"
    End If
    
    Case 2 'Incluir Tipos de Formularios
     sDate = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"
'                    "FOR_CODIGO, "
    sSQL = "INSERT INTO EXP_TAB_FORMULARIO(" & _
                    " FOR_NOME_EDITOR ," & _
                    " FOR_NOME_FORM ," & _
                    " FOR_DESCRICAO ," & _
                    " FOR_GRUPO ," & _
                    " FOR_ACESSO ," & _
                    " FOR_USU," & _
                    " FOR_DTA ," & _
                    " FOR_DTI "
    sSQL = sSQL & ") VALUES (" & _
                    cFields(2) & ", " & _
                    cFields(3) & ", " & _
                    cFields(4) & ", " & _
                    cFields(5) & "," & _
                    cFields(6) & "," & _
                    cFields(7) & "," & _
                    sDate & ", " & sDate & ")"

    Case 3 'Alterar Tipos de Formularios
    
     sDate = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"
'                 "  EXP_TAB_FORMULARIO.FOR_CODIGO = " & cFields(1)
    sSQL = "UPDATE EXP_TAB_FORMULARIO SET" & _
                    "  FOR_NOME_EDITOR = " & cFields(2) & _
                    " ,FOR_NOME_FORM = " & cFields(3) & _
                    " ,FOR_DESCRICAO = " & cFields(4) & _
                    " ,FOR_GRUPO = " & cFields(5) & _
                    " ,FOR_ACESSO = " & cFields(6) & _
                    " ,FOR_USU= " & cFields(7) & _
                    " ,FOR_DTA = " & sDate & _
                    " WHERE FOR_CODIGO = " & cFields(1) & "" & _
                    " "
    
    Case 4 'Excluir Tipos de Formularios
    sSQL = "DELETE  FROM EXP_TAB_FORMULARIO" & _
                 "  WHERE EXP_TAB_FORMULARIO.FOR_CODIGO = " & cFields(1) & _
                 " AND EXP_TAB_FORMULARIO.FOR_DTA = " & cFields(8)

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

If nAction = 2 Then  'Receber o ultimo registro que acabou dee Incluir
   sSQL = "SELECT MAX(EXP_TAB_FORMULARIO.FOR_CODIGO) AS ultimoDeFOR_CODIGO" & _
                 " FROM EXP_TAB_FORMULARIO;"
   cCONECCAO.CursorLocation = adUseClientBatch
   rs.Open sSQL, cCONECCAO
End If


cCONECCAO.CommitTrans: cCONECCAO.BeginTrans
Set TipoForm_dados = rs
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

