VERSION 5.00
Begin VB.Form frmGIFMntAberturaInventário 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abertura de inventário"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7905
   Begin VB.Frame Frame3 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   7695
      Begin VB.Frame Frame6 
         Caption         =   "Confirme seu usuário para habilitar fechamento"
         Height          =   2325
         Left            =   90
         TabIndex        =   9
         Top             =   180
         Width           =   3555
         Begin VB.ComboBox Cbo_Usuario 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            ItemData        =   "frmGIFMntAberturaInventário.frx":0000
            Left            =   90
            List            =   "frmGIFMntAberturaInventário.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   660
            Width           =   3255
         End
         Begin VB.TextBox TXTSENHA 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   555
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   9
            PasswordChar    =   "*"
            TabIndex        =   10
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Selecione Usuário"
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
            Left            =   420
            TabIndex        =   13
            Top             =   270
            Width           =   2205
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Senha"
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
            Left            =   1170
            TabIndex        =   12
            Top             =   1260
            Width           =   795
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Informações"
         Height          =   4275
         Left            =   4140
         TabIndex        =   4
         Top             =   180
         Width           =   3255
         Begin VB.Frame Frame2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Da ultima Abertura"
            Height          =   2025
            Left            =   150
            TabIndex        =   7
            Top             =   240
            Width           =   2895
            Begin VB.Label lbl_abertura_anterior 
               BackColor       =   &H0080C0FF&
               Caption         =   "Aguardando rotina para fechamento."
               Height          =   1605
               Left            =   90
               TabIndex        =   8
               Top             =   270
               Width           =   2685
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H0080FF80&
            Caption         =   "De abertura"
            Height          =   1755
            Left            =   180
            TabIndex        =   5
            Top             =   2400
            Width           =   2895
            Begin VB.Label lbl_nova_abertura 
               BackColor       =   &H0080FF80&
               Height          =   1425
               Left            =   180
               TabIndex        =   6
               Top             =   270
               Width           =   2565
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Abertura do Inventário/Confirmação"
         Height          =   1785
         Left            =   90
         TabIndex        =   2
         Top             =   2610
         Width           =   3555
         Begin VB.CommandButton cmd_Confirmar 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Confirmar Abertura"
            Enabled         =   0   'False
            Height          =   525
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Confirmar a alteração do tipo da caixa."
            Top             =   1050
            Width           =   3255
         End
         Begin VB.TextBox TXT_DT_ABERTURA 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   840
            MaxLength       =   10
            TabIndex        =   3
            Text            =   "01'01'2009"
            Top             =   390
            Width           =   1725
         End
      End
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   6510
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4770
      Width           =   1275
   End
End
Attribute VB_Name = "frmGIFMntAberturaInventário"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Private Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Private cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public Confirma_Mudanca As Boolean 'Servirá para confirmar as mudanças de alteracoes dos campos na tela
Private nxv As Integer 'NUMERO DE TENTATIVAS PARA ACERTAR A DIGITACAO DA SENHA

Private Sub cmd_Confirmar_Click()
Call Abertuta_Inventario
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
Me.TXT_DT_ABERTURA.Text = Format(Now(), "DD/MM/YYYY")
If Me.cbo_usuario.ListCount = 0 Then
   Unload Me
End If
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
Call Carrega_Usuario
Call Carrega_Ultima_Abertura

End Sub

Public Function CCTemp() As neManutencao
     Set CCTemp = New neManutencao
End Function

Private Sub Cbo_Usuario_Change()
Me.cmd_Confirmar.Enabled = False
Me.TXTSENHA.Text = ""
Me.TXTSENHA.Enabled = True
Me.TXTSENHA.SetFocus
End Sub

Private Sub Cbo_Usuario_Click()
Me.cmd_Confirmar.Enabled = False
Me.TXTSENHA.Text = ""
Me.TXTSENHA.Enabled = True
Me.TXTSENHA.SetFocus
End Sub

Private Sub Cbo_Usuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Len(sNomeBanco) > 0 Then
      Me.TXTSENHA.Enabled = True
      Me.TXTSENHA.SetFocus
   End If
End If

End Sub

Public Function Carrega_Usuario()
Dim cRecAux As ADODB.Recordset
Dim nx As Integer

On Error GoTo Erro
Me.MousePointer = vbHourglass

Set cRecAux = CCTempneUsuario.USUARIO_Consultar(sNomeBanco)

If cRecAux Is Nothing Then
   MsgBox "Não Existem Usuarios cadastrados!"
   Exit Function
   
ElseIf cRecAux.RecordCount = 0 Then
   MsgBox "Não Existem Usuarios cadastrados ou habilitados!"
   Exit Function
End If

cRecAux.MoveFirst

If cRecAux.RecordCount > 0 Then
   cbo_usuario.Clear
   For nx = 1 To cRecAux.RecordCount
       cbo_usuario.AddItem cRecAux!USU_USUARIO
       cbo_usuario.ItemData(nx - 1) = cRecAux!USU_CODIGO
       cRecAux.MoveNext
   Next
Else
   MsgBox "Não Existem usuarios cadastrados!"
End If

Me.MousePointer = vbDefault
Set cRecAux = Nothing

Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Function

Private Sub TXTSENHA_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If Confirmar_Senha Then
      Call Dados_Banco
      Exit Sub
   Else
      MsgBox "Senha inválida, Tente novamente!"
   End If
End If

If KeyAscii = 27 Then
   Me.cbo_usuario.SetFocus
End If

End Sub

Private Function Confirmar_Senha() As Boolean
Dim nx As Integer

On Error GoTo Erro


If Me.cbo_usuario.ListIndex = -1 Then
   Exit Function
End If

Confirmar_Senha = False
Me.MousePointer = vbHourglass

Set cRec = New ADODB.Recordset
Set cRec = CCTempneUsuario.USUARIO_Consultar(sNomeBanco, Me.cbo_usuario.ItemData(Me.cbo_usuario.ListIndex))

If cRec.RecordCount > 0 Then
    cRec.MoveFirst
    For nx = 0 To cRec.RecordCount - 1
        If UnCripta(Trim(cRec!USU_SENHA)) = Trim(Me.TXTSENHA.Text) Then
           Confirmar_Senha = True
           Exit For
        End If
        cRec.MoveNext
    Next
End If
Me.MousePointer = vbDefault
Set cRec = Nothing

Exit Function

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault
Set cRec = Nothing

End Function

Private Function Dados_Banco()
Dim nx As Integer

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set cRec = New ADODB.Recordset
Set cRec = CCTemp.MANUTENCAO_INVENTARIO_Qt_Registro(sNomeBanco)

If cRec.RecordCount > 0 Then
    cRec.MoveFirst
    Me.lbl_nova_abertura.Caption = "Neste momento o banco contém " & _
                                   Format(cRec!TOTAL_REG + cRec!INVENTARIADO, "#,###.###") & _
                                   " Registros dos quais apenas, " & _
                                   Format(cRec!INVENTARIADO, "###,###") & " serão atualizados."
    Me.cmd_Confirmar.Enabled = True
    Me.cmd_Confirmar.SetFocus
'    MsgBox "Senha válida, pode realizar o Fechamento!"
Else
    Me.lbl_nova_abertura.Caption = "Neste momento o banco não contém registros." & _
                                   "Não será realizado nenhuma abertura."

End If
Me.MousePointer = vbDefault
Set cRec = Nothing

Exit Function

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault
Set cRec = Nothing

End Function
Private Function Carrega_Ultima_Abertura()
Dim nx As Integer

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set cRec = New ADODB.Recordset
Set cRec = CCTempneLog.LOG_CONSULTAR_Ult_Abertura_Inv(sNomeBanco)

If cRec.RecordCount > 0 Then
    cRec.MoveFirst
    Me.lbl_abertura_anterior.Caption = cRec!LOG_OBSERVACAO
Else
    Me.lbl_abertura_anterior.Caption = "Neste momento o banco não contém registros." & _
                                       "Será realizado a primeira abertura."

End If

Me.MousePointer = vbDefault
Set cRec = Nothing

Exit Function

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault
Set cRec = Nothing

End Function
Private Function Abertuta_Inventario()
Dim nx As Integer

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set cRec = New ADODB.Recordset

Set cRec = CCTemp.MANUTENCAO_INVENTARIO_Abertura(sNomeBanco, _
                                                 Me.cbo_usuario.List(Me.cbo_usuario.ListIndex), _
                                                 Me.TXT_DT_ABERTURA.Text)
                                                 
MsgBox "Fechamento realizado com Sucesso!"
Me.cmd_Confirmar.Enabled = False
Me.TXTSENHA.Text = ""
Me.lbl_nova_abertura = ""
Call Carrega_Ultima_Abertura
Me.cbo_usuario.ListIndex = -1
Me.cbo_usuario.SetFocus


Me.MousePointer = vbDefault
Set cRec = Nothing

Exit Function

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault
Set cRec = Nothing

End Function


