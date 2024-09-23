VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmGIFMntImportaMovInventário 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação do movimento de inventário"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   9765
   Begin VB.CommandButton cmd_atualiza 
      BackColor       =   &H0080FF80&
      Caption         =   "Confirme Atualização"
      Height          =   405
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   510
      Width           =   2055
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   8430
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1530
      Width           =   1275
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   1470
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1530
      Width           =   1005
   End
   Begin VB.TextBox txt_atualizado 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   4050
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1530
      Width           =   1005
   End
   Begin VB.TextBox txt_atualizadas 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   6510
      MaxLength       =   6
      TabIndex        =   0
      Top             =   1530
      Width           =   1005
   End
   Begin ComctlLib.ProgressBar PBar1 
      Height          =   225
      Left            =   150
      TabIndex        =   5
      Top             =   1140
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
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
      Left            =   2100
      TabIndex        =   10
      Top             =   90
      Width           =   7665
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
      Left            =   180
      TabIndex        =   9
      Top             =   90
      Width           =   1905
   End
   Begin VB.Label Label5 
      Caption         =   "Total encontrado : "
      Height          =   285
      Left            =   150
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Total lidas : "
      Height          =   285
      Left            =   2850
      TabIndex        =   7
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label Label8 
      Caption         =   "Atualizadas : "
      Height          =   285
      Left            =   5310
      TabIndex        =   6
      Top             =   1560
      Width           =   1185
   End
End
Attribute VB_Name = "frmGIFMntImportaMovInventário"
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

Nada = Me.lbl_arquivo.Caption

If Dir$(Nada) = "" Then
   MsgBox "Arquivo de importação não encontrado, Procure o responsável! " & Nada, 16, "Atualização Cancelada"
   Exit Sub
End If

RESPOSTA = MsgBox("Importar dados do Inventário?", 20, "Sim/Não?")

On Error Resume Next

If RESPOSTA = 6 Then
   Call CCTemp.MANUTENCAO_INVENTARIO_Importar(sNomeBanco, sNomeUsuario, Format(Now(), "dd/mm/yyyy"), Nada)
   Me.MousePointer = vbHourglass
   MsgBox "Importação realizada com Sucesso!"
      
End If

Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub

Private Sub cmdfechar_Click()
Unload Me
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
Me.Top = 0
Me.Left = 0
End Sub

Public Function CCTemp() As neManutencao
     Set CCTemp = New neManutencao
End Function
