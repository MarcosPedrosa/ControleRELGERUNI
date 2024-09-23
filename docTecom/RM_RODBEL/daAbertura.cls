VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "daAbertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI A CAMADA DE ABERTURA BANCO DADOS
' ***********************************************************
' ***********************************************************
' ***********************************************************
Option Explicit

Public Function Coneccao_ODBC(ByVal sNomeBanco As String, _
                              ByVal Tipo_Abertura As String) As ADODB.Connection
On Error GoTo Erro
Dim CONEXAO As Connection
Dim ADOConnection As ADODB.Connection
Dim sFileDSN As String

Set ADOConnection = New ADODB.Connection

If Tipo_Abertura = "A" Then
   sFileDSN = Trim("FileDSN=" & sNomeBanco)
   ADOConnection.Open sFileDSN
   ADOConnection.BeginTrans
   Set CONEXAO = ADOConnection
End If

Exit Function

Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function Coneccao_OLEDB(ByVal sNomeBanco As String, _
                               ByVal Tipo_Abertura As String) As ADODB.Connection
On Error GoTo Erro
Dim CONEXAO As Connection
Dim ADOConnection As ADODB.Connection
Dim ssql As String

Set ADOConnection = New ADODB.Connection

    
Rem sSQL = "Provider=SQLOLEDB.1;Persist Security Info=True;Data Source=10.0.0.3;Initial Catalog=teklogix;User ID=etiquetas;Password=etiquetas;"
Rem ssql = "Driver={Microsoft ODBC for Oracle};Server=XE;uid=Default_Acesso;pwd=Default;"

Set ADOConnection = New ADODB.Connection
ADOConnection.Open sNomeBanco
ADOConnection.BeginTrans
Set Coneccao_OLEDB = ADOConnection
        
Exit Function

Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

