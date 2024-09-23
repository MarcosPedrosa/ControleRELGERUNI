VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmGIFDisImportaMovInventario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação dados inventário"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   9885
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3210
      TabIndex        =   20
      Top             =   2340
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   2340
      Width           =   1335
   End
   Begin VB.TextBox txt_atualizadas 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   17
      Top             =   1560
      Width           =   1005
   End
   Begin VB.TextBox txt_atualizado 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   4020
      MaxLength       =   6
      TabIndex        =   15
      Top             =   1560
      Width           =   1005
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   13
      Top             =   1590
      Width           =   1005
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1560
      Width           =   1275
   End
   Begin VB.CommandButton cmd_atualiza 
      Caption         =   "Confirme Atualização"
      Height          =   405
      Left            =   7620
      TabIndex        =   10
      Top             =   540
      Width           =   2055
   End
   Begin VB.TextBox txt_Tipo_arquivo 
      Height          =   315
      Left            =   7170
      TabIndex        =   7
      Text            =   "*.txt"
      Top             =   4230
      Width           =   2715
   End
   Begin VB.FileListBox F 
      Height          =   480
      Left            =   4710
      TabIndex        =   3
      Top             =   3420
      Width           =   4995
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   60
      TabIndex        =   1
      Top             =   3450
      Width           =   4455
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   4620
      Width           =   825
   End
   Begin ComctlLib.ProgressBar PBar1 
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   1170
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label8 
      Caption         =   "Atualizadas : "
      Height          =   285
      Left            =   5280
      TabIndex        =   18
      Top             =   1590
      Width           =   1185
   End
   Begin VB.Label Label7 
      Caption         =   "Total lidas : "
      Height          =   285
      Left            =   2820
      TabIndex        =   16
      Top             =   1590
      Width           =   1185
   End
   Begin VB.Label Label5 
      Caption         =   "Total encontrado : "
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Arquivo selecionado.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   9
      Top             =   120
      Width           =   1905
   End
   Begin VB.Label lbl_arquivo 
      Caption         =   "C:\invcoletor.txt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2070
      TabIndex        =   8
      Top             =   120
      Width           =   7665
   End
   Begin VB.Label Label4 
      Caption         =   "Digite o tipo do arquivo.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4890
      TabIndex        =   6
      Top             =   4260
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Selecione o arquivo desejado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4710
      TabIndex        =   5
      Top             =   3090
      Width           =   2565
   End
   Begin VB.Label Label2 
      Caption         =   "Localize o diretório onde se encontra o arquivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3150
      Width           =   4365
   End
   Begin VB.Label Label1 
      Caption         =   "Selecione o drive.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   4350
      Width           =   1785
   End
End
Attribute VB_Name = "frmGIFDisImportaMovInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Private Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Private cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public Confirma_Mudanca As Boolean 'Servirá para confirmar as mudanças de alteracoes dos campos na tela

Private Sub cmd_atualiza_Click()
Dim nx As Integer
Dim Nada As String
Dim RESPOSTA As Integer
Dim CONTA As Double
Dim x As Double
Dim Y As Double
Dim ADOConnection As ADODB.Connection
Dim cConect As daAbertura
Dim sSql As String
Dim sPesquisa As String
Dim rs As ADODB.Recordset

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set ADOConnection = cConect.Coneccao(sNomeBanco, "A")

Nada = Me.lbl_arquivo.Caption

If Dir$(Nada) = "" Then
   MsgBox "Arquivo de importação não encontrado, Procure o responsável! " & Nada, 16, "Atualização Cancelada"
   Exit Sub
End If

Close #11
Open Nada For Random Access Read Write As #11 Len = Len(Arq_Mov_Inventario)

Y = LOF(11) / Len(Arq_Mov_Inventario)
Me.txtlidos.Text = Y

RESPOSTA = MsgBox("Ler dados do Inventário?", 20, "Sim/Não?")

sSql = "DELETE FROM EXP_TMP_INVENTARIO "

Set rs = New ADODB.Recordset
ADOConnection.CursorLocation = adUseClientBatch
rs.Open sSql, ADOConnection

On Error Resume Next

If RESPOSTA = 6 Then
      nAtualizadas = 0
      PBar1.Value = CONTA
      PBar1.Visible = True
      PBar1.Min = 0
      Y = LOF(11) / Len(Arq_Mov_Inventario)
      PBar1.Max = Y
      x = 0
      CONTA = 0
      For Y = 1 To LOF(11) / Len(Arq_Mov_Inventario)
          CONTA = CONTA + 1
          PBar1.Value = CONTA
          Get 11, Y, Arq_Mov_Inventario
          nAtualizadas = nAtualizadas + 1
          sSql = "INSERT INTO EXP_TMP_INVENTARIO (ID_ETIQUETA,ID_BORDERO) VALUES ('" & _
                 Trim(Arq_Mov_Inventario.FRegistro) & "','" & _
                 Trim(Arq_Mov_Inventario.FTipo) & "')"
          ADOConnection.CursorLocation = adUseClientBatch
          rs.Open sSql, ADOConnection
          Me.txt_atualizado.Text = CONTA
          Me.txt_atualizadas.Text = nAtualizadas
          Me.txt_atualizado.Refresh
          Me.txt_atualizadas.Refresh
      Next
      
      sSql = "UPDATE ETI " & _
             "SET ETI.ID_BORDERO = INV.ID_BORDERO " & _
             "FROM EXP_TMP_INVENTARIO INV " & _
             "INNER JOIN ETIQUETA ETI " & _
             "ON INV.ID_ETIQUETA = ETI.ID_ETIQUETA AND INV.ID_BORDERO <> ETI.ID_BORDERO"

      Set rs = New ADODB.Recordset
      ADOConnection.CursorLocation = adUseClientBatch
      rs.Open sSql, ADOConnection
      
      ADOConnection.CommitTrans
      MsgBox "Atualização realizada com Sucesso!"
End If
Close #11
Set ADOConnection = Nothing
Set cConect = Nothing

''''Dim nx As Integer
''''Dim Nada As String
''''Dim RESPOSTA As Integer
''''Dim CONTA As Double
''''Dim x As Double
''''Dim Y As Double
''''
''''Nada = Me.lbl_arquivo.Caption
''''
''''If Dir$(Nada) = "" Then
''''   MsgBox "Arquivo de importação não encontrado, Procure o responsável! " & Nada, 16, "Atualização Cancelada"
''''   Exit Sub
''''End If
''''
''''Close #11
''''Open Nada For Random Access Read Write As #11 Len = Len(Arq_Mov_Inventario)
''''
''''Y = LOF(11) / Len(Arq_Mov_Inventario)
''''Me.txtlidos.Text = Y
''''
''''RESPOSTA = MsgBox("Ler dados do Inventário?", 20, "Sim/Não?")
''''
''''On Error Resume Next
''''
''''If RESPOSTA = 6 Then
''''      nAtualizadas = 0
''''      PBar1.Value = CONTA
''''      PBar1.Visible = True
''''      PBar1.Min = 0
''''      Y = LOF(11) / Len(Arq_Mov_Inventario)
''''      PBar1.Max = Y
''''      x = 0
''''      CONTA = 0
''''      For Y = 1 To LOF(11) / Len(Arq_Mov_Inventario)
''''        CONTA = CONTA + 1
''''        PBar1.Value = CONTA
''''        Get 11, Y, Arq_Mov_Inventario
''''        Call CCTemp.EXPEDICAO_Atualiza_Inventario(sNomeBanco, _
''''                                                Trim(Arq_Mov_Inventario.FRegistro), _
''''                                                Trim(Arq_Mov_Inventario.FTipo))
''''        Me.txt_atualizado.Text = CONTA
''''        Me.txt_atualizadas.Text = nAtualizadas
''''        Me.txt_atualizado.Refresh
''''        Me.txt_atualizadas.Refresh
''''      Next
''''      MsgBox "Atualização realizada com Sucesso!"
''''Else
''''     Exit Sub
''''End If
''''Close #11

End Sub

Private Sub cmdfechar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim nx As Integer
Dim Nada As String
Dim RESPOSTA As Integer
Dim CONTA As Double
Dim x As Double
Dim Y As Double
Dim ADOConnection As ADODB.Connection
Dim cConect As daAbertura
Dim sSql As String
Dim sPesquisa As String
Dim rs As ADODB.Recordset

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set ADOConnection = cConect.Coneccao(sNomeBanco, "A")

Nada = Me.lbl_arquivo.Caption

If Dir$(Nada) = "" Then
   MsgBox "Arquivo de importação não encontrado, Procure o responsável! " & Nada, 16, "Atualização Cancelada"
   Exit Sub
End If

Close #11
Open Nada For Random Access Read Write As #11 Len = Len(Arq_Mov_Inventario)

Y = LOF(11) / Len(Arq_Mov_Inventario)
Me.txtlidos.Text = Y

RESPOSTA = MsgBox("Ler dados do Inventário?", 20, "Sim/Não?")

On Error Resume Next

If RESPOSTA = 6 Then
      nAtualizadas = 0
      PBar1.Value = CONTA
      PBar1.Visible = True
      PBar1.Min = 0
      Y = LOF(11) / Len(Arq_Mov_Inventario)
      PBar1.Max = Y
      x = 0
      CONTA = 0
      For Y = 1 To LOF(11) / Len(Arq_Mov_Inventario)
          CONTA = CONTA + 1
          PBar1.Value = CONTA
          Get 11, Y, Arq_Mov_Inventario
          
          If CONTA = 1100 Then
             CONTA = 1100
          End If
          
          sSql = "SELECT ID_ETIQUETA,ID_BORDERO FROM ETIQUETA " & _
                 " WHERE ID_ETIQUETA = '" & Trim(Arq_Mov_Inventario.FRegistro) & "' " & _
                 " AND   ID_BORDERO <> '" & Trim(Arq_Mov_Inventario.FTipo) & "'"

          Set rs = New ADODB.Recordset
          ADOConnection.CursorLocation = adUseClientBatch
          rs.Open sSql, ADOConnection

          
          If rs.RecordCount > 0 Then
                nAtualizadas = nAtualizadas + 1
                sSql = "UPDATE ETIQUETA SET " & _
                       "ID_BORDERO = '" & Trim(Arq_Mov_Inventario.FTipo) & "'" & _
                       " WHERE ID_ETIQUETA = '" & Trim(Arq_Mov_Inventario.FRegistro) & "'"
                ADOConnection.CursorLocation = adUseClientBatch
                rs.Open sSql, ADOConnection
          End If
        
          Me.txt_atualizado.Text = CONTA
          Me.txt_atualizadas.Text = nAtualizadas
          Me.txt_atualizado.Refresh
          Me.txt_atualizadas.Refresh
      Next
      ADOConnection.CommitTrans
      Set ADOConnection = Nothing
      Set cConect = Nothing
      MsgBox "Atualização realizada com Sucesso!"
Else
     Exit Sub
End If
Close #11
End Sub

Private Sub Command2_Click()
Dim nx As Integer
Dim Nada As String
Dim RESPOSTA As Integer
Dim CONTA As Double
Dim x As Double
Dim Y As Double
Dim ADOConnection As ADODB.Connection
Dim cConect As daAbertura
Dim sSql As String
Dim sPesquisa As String
Dim rs As ADODB.Recordset

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set ADOConnection = cConect.Coneccao(sNomeBanco, "A")

Nada = Me.lbl_arquivo.Caption

If Dir$(Nada) = "" Then
   MsgBox "Arquivo de importação não encontrado, Procure o responsável! " & Nada, 16, "Atualização Cancelada"
   Exit Sub
End If

Close #11
Open Nada For Random Access Read Write As #11 Len = Len(Arq_Mov_Inventario)

Y = LOF(11) / Len(Arq_Mov_Inventario)
Me.txtlidos.Text = Y

RESPOSTA = MsgBox("Ler dados do Inventário?", 20, "Sim/Não?")

sSql = "DELETE FROM EXP_TMP_INVENTARIO "

Set rs = New ADODB.Recordset
ADOConnection.CursorLocation = adUseClientBatch
rs.Open sSql, ADOConnection

On Error Resume Next

If RESPOSTA = 6 Then
      nAtualizadas = 0
      PBar1.Value = CONTA
      PBar1.Visible = True
      PBar1.Min = 0
      Y = LOF(11) / Len(Arq_Mov_Inventario)
      PBar1.Max = Y
      x = 0
      CONTA = 0
      For Y = 1 To LOF(11) / Len(Arq_Mov_Inventario)
          CONTA = CONTA + 1
          PBar1.Value = CONTA
          Get 11, Y, Arq_Mov_Inventario
          nAtualizadas = nAtualizadas + 1
          sSql = "INSERT INTO EXP_TMP_INVENTARIO (ID_ETIQUETA,ID_BORDERO) VALUES ('" & _
                 Trim(Arq_Mov_Inventario.FRegistro) & "','" & _
                 Trim(Arq_Mov_Inventario.FTipo) & "')"
          ADOConnection.CursorLocation = adUseClientBatch
          rs.Open sSql, ADOConnection
          Me.txt_atualizado.Text = CONTA
          Me.txt_atualizadas.Text = nAtualizadas
          Me.txt_atualizado.Refresh
          Me.txt_atualizadas.Refresh
      Next
      
      sSql = "UPDATE ETI " & _
             "SET ETI.ID_BORDERO = INV.ID_BORDERO " & _
             "FROM EXP_TMP_INVENTARIO INV " & _
             "INNER JOIN ETIQUETA ETI " & _
             "ON INV.ID_ETIQUETA = ETI.ID_ETIQUETA AND INV.ID_BORDERO <> ETI.ID_BORDERO"

      Set rs = New ADODB.Recordset
      ADOConnection.CursorLocation = adUseClientBatch
      rs.Open sSql, ADOConnection
      
      ADOConnection.CommitTrans
      Set ADOConnection = Nothing
      Set cConect = Nothing
      MsgBox "Atualização realizada com Sucesso!"
Else
     Exit Sub
End If
Close #11

End Sub

Private Sub Form_Activate()
If Flag_ativo = True Then
   Exit Sub
End If
Me.Top = 0
Me.Left = 0
Flag_ativo = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' para funcionar , tem que mudar o keyPreviwe=true
If KeyCode = 13 Then
      SendKeys "{TAB}"
ElseIf KeyCode = 27 Then
'   If Me.ActiveControl.TabIndex < 8 Then
'      If Me.CMD_SALVAR.Enabled = True Then
'        If 6 = MsgBox("Deseja realmente sair deste módulo?", 32 + 4) Then
'           Unload Me
'        End If
'      Else
        Unload Me
'      End If
'   Else
'       SendKeys "+{TAB}" ' retornar campo
'   End If
End If


End Sub

Private Sub Form_Load()
Dim nx As Integer

Me.Top = 0
Me.Left = 0
End Sub
Private Sub Drive1_Change()

On Error GoTo erro
Dir1.Path = Mid$(Me.Drive1.List(Me.Drive1.ListIndex), 1, 2) & "\"
Exit Sub
erro:

If Err.Number = 68 Then
   MsgBox "Caminho inválido, escolha outro!"
End If

End Sub
'
'Private Sub txt_arquivo_Change()
'Me.File1.Parent = Me.txt_arquivo.Text
'End Sub

Public Function CCTemp() As neExpedicao
     Set CCTemp = New neExpedicao
End Function

