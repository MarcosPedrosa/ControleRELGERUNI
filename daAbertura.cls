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

Public Function Coneccao(ByVal sNameBanco11 As String, ByVal Tipo_Abertura As String) As Connection
On Error GoTo Erro

Dim CConnection As Connection
Dim sFileDSN As String

Dim sFileUSU As String
Dim sFileSEN As String
Dim nTamUSU, nTamSEN, nInicio, nFinal As Integer

Set CConnection = New Connection

 If Tipo_Abertura = "A" Then
    If Not ADOConnection Is Nothing Then
       If ADOConnection.State = 1 Then
          If ADOConnection.Errors.Count = 1 Then
             ADOConnection.Errors.Clear
             sFileDSN = Trim(sNameBanco1)
             ADOConnection.BeginTrans
          End If
          Set Coneccao = ADOConnection
          Exit Function
       End If
    End If
    
    nInicio = InStr(1, sNameBanco1, "Catalog=") + 8
    nFinal = InStr(1, sNameBanco1, ";User ID=")
    sFileDSN = Trim(Mid$(Trim(sNameBanco1), nInicio, nFinal - nInicio))
    nTamUSU = InStr(1, sNameBanco1, ",") + 1
    
    nInicio = InStr(1, sNameBanco1, "User ID=") + 8
    nFinal = InStr(1, sNameBanco1, ";Password=")
    sFileUSU = Trim(Mid$(Trim(sNameBanco1), nInicio, nFinal - nInicio))
    nTamSEN = InStr(nTamUSU + 1, sNameBanco1, ",") + 1
    
    nInicio = InStr(1, sNameBanco1, "Password=") + 9
    nFinal = Len(sNameBanco1)
    sFileSEN = Trim(Mid$(Trim(sNameBanco1), nInicio, nFinal - nInicio))
    
    CConnection.Open sFileDSN, sFileUSU, sFileSEN

'    sFileDSN = Trim(sNameBanco1)
'    CConnection.Open sFileDSN
    CConnection.BeginTrans
    Set Coneccao = CConnection
Else
    If Tipo_Abertura = "F" And Not ADOConnection Is Nothing Then
       ADOConnection.Close
       Set ADOConnection = Nothing
    End If
End If

Exit Function

Erro:

If Err.Description = "Não é possível iniciar mais transações nesta sessão." Then Exit Function

Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function Coneccao_OLEDB(ByVal sBanco11 As String, _
                               ByVal Tipo_Abertura As String) As ADODB.Connection
On Error GoTo Erro
Dim CONEXAO As Connection
Dim sSQL As String

Dim sFileUSU As String
Dim sFileSEN As String
Dim nTamUSU, nTamSEN, nInicio, nFinal As Integer

Set CONEXAO = New ADODB.Connection

'sFileDSN = Mid$(Trim(sBanco1), 1, InStr(1, sBanco1, ",") - 1)
'nTamUSU = InStr(1, sBanco1, ",") + 1
'
'sFileUSU = Trim(Mid$(Trim(sBanco1), nTamUSU, InStr(nTamUSU, sBanco1, ",") - nTamUSU))
'nTamSEN = InStr(nTamUSU + 1, sBanco1, ",") + 1
'
'sFileSEN = Trim(Mid$(Trim(sBanco1), nTamSEN, Len(Trim(sBanco1))))
nInicio = InStr(1, sBanco11, "Catalog=") + 8
nFinal = InStr(1, sBanco11, ";User ID=")
sFileDSN = Trim(Mid$(Trim(sBanco11), nInicio, nFinal - nInicio))
nTamUSU = InStr(1, sBanco11, ",") + 1

nInicio = InStr(1, sBanco11, "User ID=") + 8
nFinal = InStr(1, sBanco11, ";Password=")
sFileUSU = Trim(Mid$(Trim(sBanco11), nInicio, nFinal - nInicio))
nTamSEN = InStr(nTamUSU + 1, sBanco11, ",") + 1

nInicio = InStr(1, sBanco11, "Password=") + 9
nFinal = Len(sBanco11)
sFileSEN = Trim(Mid$(Trim(sBanco11), nInicio, nFinal - nInicio))

CONEXAO.Open sFileDSN, sFileUSU, sFileSEN

'CONEXAO.Open sFileDSN, sFileUSU, sFileSEN

'CONEXAO.Open sBanco1
CONEXAO.BeginTrans
Set Coneccao_OLEDB = CONEXAO
        
Exit Function

Erro:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


''''Public Function Coneccao(ByVal sBanco11 As String, ByVal Tipo_Abertura As String) As Connection
''''On Error GoTo Erro
''''
''''Dim CConnection As Connection
''''Dim sFileDSN As String
''''
''''Set CConnection = New Connection
''''
'''' If Tipo_Abertura = "A" Then
''''    If Not ADOConnection Is Nothing Then
''''       If ADOConnection.State = 1 Then
''''          If ADOConnection.Errors.Count = 1 Then
''''             ADOConnection.Errors.Clear
''''             sFileDSN = Trim(sBanco11)
''''             ADOConnection.BeginTrans
''''          End If
''''          Set Coneccao = ADOConnection
''''          Exit Function
''''       End If
''''    End If
''''
''''    sFileDSN = Trim(sBanco11)
''''    CConnection.Open sFileDSN
''''    CConnection.BeginTrans
''''    Set Coneccao = CConnection
''''Else
''''    If Tipo_Abertura = "F" And Not ADOConnection Is Nothing Then
''''       ADOConnection.Close
''''       Set ADOConnection = Nothing
''''    End If
''''End If
''''
''''Exit Function
''''
''''Erro:
''''
''''If Err.Description = "Não é possível iniciar mais transações nesta sessão." Then Exit Function
''''
''''Err.Raise Err.Number, Err.Source, Err.Description
''''
''''End Function

'Public Function ConeccaoRM(ByVal sBanco11 As String, ByVal Tipo_Abertura As String) As Connection
'On Error GoTo Erro
'
'Dim CConnection As Connection
'Dim sFileDSN As String
'
'Set CConnection = New Connection
'
'If Tipo_Abertura = "A" Then
'   If Not ADOConnection Is Nothing Then
'      If ADOConnection.State = 1 Then
'         If ADOConnection.Errors.Count = 1 Then
'            ADOConnection.Errors.Clear
'            sFileDSN = Trim(sBanco11)
'            ADOConnection.BeginTrans
'         End If
'         Set ConeccaoRM = ADOConnection
'         Exit Function
'      End If
'   End If
'
'   sFileDSN = Trim(sBanco11)
'   CConnection.Open sFileDSN
'   CConnection.BeginTrans
'   Set ConeccaoRM = CConnection
'Else
'   If Tipo_Abertura = "F" And Not ADOConnection Is Nothing Then
'      ADOConnection.Close
'      Set ADOConnection = Nothing
'   End If
'End If
'
'Exit Function
'
'Erro:
'
'If Err.Description = "Não é possível iniciar mais transações nesta sessão." Then Exit Function
'
'Err.Raise Err.Number, Err.Source, Err.Description
'
'End Function



