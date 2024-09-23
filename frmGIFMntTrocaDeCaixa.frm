VERSION 5.00
Begin VB.Form frmGIFMntTrocaDeCaixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Troca do Tipo da caixa"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmGIFMntTrocaDeCaixa.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   9405
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   8070
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Fechar a Tela"
      Top             =   2670
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecione o novo tipo da caixa "
      Height          =   1185
      Left            =   90
      TabIndex        =   8
      Top             =   1380
      Width           =   9225
      Begin VB.CommandButton cmd_Confirmar 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Confirmar"
         Height          =   375
         Left            =   7980
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Confirmar a alteração do tipo da caixa."
         Top             =   480
         Width           =   1155
      End
      Begin VB.ComboBox cbo_tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3270
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   3195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo caixa.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1830
         TabIndex        =   9
         Top             =   510
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selecione a caixa Origem"
      Height          =   1185
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   9225
      Begin VB.TextBox TXT_TIPO_CX 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5190
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   390
         Width           =   3825
      End
      Begin VB.CommandButton cmd_confirmar_CX 
         BackColor       =   &H00C0FFC0&
         Height          =   405
         Left            =   4470
         Picture         =   "frmGIFMntTrocaDeCaixa.frx":03F6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Validar o numero da Caixa digitada"
         Top             =   450
         Width           =   435
      End
      Begin VB.TextBox TXT_CX 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2130
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "Digite o número de uma caixa para podeer alterar o tipo de caixa"
         Top             =   420
         Width           =   2805
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6960
         TabIndex        =   10
         Top             =   150
         Width           =   390
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nº CAIXA.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   750
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmGIFMntTrocaDeCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Private Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Private cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public Confirma_Mudanca As Boolean 'Servirá para confirmar as mudanças de alteracoes dos campos na tela

Private Sub cbo_tipo_Click()
Me.cmd_Confirmar.Enabled = True
End Sub

Private Sub cmd_Confirmar_Click()

On Error GoTo Erro

Set cRec = New ADODB.Recordset

If Len(Trim(Me.TXT_CX.Text)) = 0 Then
   MsgBox "Digite o número de uma caixa para ser validada!"
   Me.cmd_Confirmar.Enabled = False
   Me.cbo_tipo.ListIndex = -1
   Me.cbo_tipo.Enabled = False
   Me.TXT_TIPO_CX.Text = ""
   Me.TXT_CX.Text = ""
   Me.TXT_CX.SetFocus
   Exit Sub
End If

If Me.cbo_tipo.ListIndex = -1 Then
   MsgBox "Selecione o tipo da caixa para ser validada!"
   Me.cmd_Confirmar.Enabled = False
   Me.cbo_tipo.ListIndex = -1
   Me.cbo_tipo.Enabled = False
   Me.TXT_TIPO_CX.Text = ""
   Me.TXT_CX.Text = ""
   Me.TXT_CX.SetFocus
   Exit Sub
End If

Me.MousePointer = vbHourglass

Set cRec = CCTemp.MANUTENCAO_CAIXA_Alterar_Tipo(sNomeBanco, Me.TXT_CX.Text, (Me.cbo_tipo.ListIndex) + 1)

MsgBox "Caixa atualizada com sucesso!"
Me.cmd_Confirmar.Enabled = False
Me.cbo_tipo.ListIndex = -1
Me.cbo_tipo.Enabled = False
Me.TXT_TIPO_CX.Text = ""
Me.TXT_CX.Text = ""
Me.TXT_CX.SetFocus

Set cRec = Nothing
Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault
Me.cmd_Confirmar.Enabled = False
Me.TXT_TIPO_CX.Text = ""
Me.TXT_CX.Text = ""
Me.TXT_CX.SetFocus

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
'Call Limpar_campos
'Call Desabilitar_Campos

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
'ElseIf Me.ActiveControl.TabIndex = 50 Then
'       If KeyCode = 35 Then
'          SendKeys "{END}" ' FIM DO campo
'       End If
End If


End Sub

Private Sub Form_Load()
Dim nx As Integer

Me.Top = 0
Me.Left = 0

Me.cbo_tipo.AddItem "IK10"
Me.cbo_tipo.AddItem "IK33"
Me.cbo_tipo.AddItem "PAPELAO"
Me.cbo_tipo.ListIndex = -1
Me.cbo_tipo.Enabled = False
Me.cmd_Confirmar.Enabled = False
End Sub

Private Sub cmd_confirmar_CX_Click()
Call Confirmar_caixa
End Sub
Function Confirmar_caixa()

On Error GoTo Erro

Set cRec = New ADODB.Recordset

If Len(Trim(Me.TXT_CX.Text)) = 0 Then
   MsgBox "Digite o número de uma caixa para ser validada!"
   Me.MousePointer = vbDefault
   Me.TXT_CX.Text = ""
   Me.TXT_CX.SetFocus
   Exit Function
End If

Me.MousePointer = vbHourglass

Set cRec = CCTemp.MANUTENCAO_CAIXA_Consultar(sNomeBanco, Me.TXT_CX.Text)

If cRec.RecordCount > 0 Then
   Me.cmd_Confirmar.Enabled = True
   cRec.MoveFirst
   Me.TXT_TIPO_CX.Text = cRec.Fields("TIPO_CAIXA")
   Me.cbo_tipo.Enabled = True
   Me.cbo_tipo.ListIndex = 0
   Me.cbo_tipo.SetFocus
Else
   MsgBox "Não existe Caixa com esta numeração, redigite!"
   Me.cmd_Confirmar.Enabled = False
   Me.TXT_TIPO_CX.Text = ""
   Me.TXT_CX.Text = ""
   Me.TXT_CX.SetFocus
End If

Set cRec = Nothing
Me.MousePointer = vbDefault

Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Function

Public Function CCTemp() As neManutencao
     Set CCTemp = New neManutencao
End Function

Private Sub TXT_CX_Change()
If Not Testa_Numerico(Me.TXT_CX.Text, Len(Me.TXT_CX.Text)) Then
   MsgBox "Só aceita numeros, redigite"
   Me.TXT_CX.Text = Mid$(Me.TXT_CX.Text, 1, Len(Trim(Me.TXT_CX.Text)) - 1)
   Me.TXT_CX.SetFocus
   SendKeys "{END}"
End If
Me.cbo_tipo.Enabled = False
Me.cbo_tipo.ListIndex = -1
Me.cmd_Confirmar.Enabled = False
End Sub

Private Sub TXT_CX_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Len(Trim(Me.TXT_CX.Text)) > 0 Then
   Call Confirmar_caixa
End If

End Sub
