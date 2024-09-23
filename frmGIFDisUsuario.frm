VERSION 5.00
Begin VB.Form frmGIFDisUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuário"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frmGIFDisUsuario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6690
   Begin VB.CommandButton cmdnovo 
      Caption         =   "&Novo"
      Height          =   330
      Left            =   2730
      TabIndex        =   17
      Top             =   2010
      Width           =   1275
   End
   Begin VB.CommandButton cmdexcluir 
      Caption         =   "&Excluir"
      Height          =   330
      Left            =   1425
      TabIndex        =   16
      Top             =   2010
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdsalvar 
      Caption         =   "&Salvar"
      Height          =   330
      Left            =   4050
      TabIndex        =   7
      Top             =   2010
      Width           =   1275
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   5355
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2010
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   90
      TabIndex        =   12
      Top             =   870
      Width           =   6525
      Begin VB.CheckBox CHK_HABILITADO 
         Caption         =   "Tem acesso ao sistema."
         Height          =   225
         Left            =   2940
         TabIndex        =   18
         Top             =   270
         Width           =   2040
      End
      Begin VB.TextBox txtConfNovasenha 
         Alignment       =   1  'Right Justify
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4290
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   570
         Width           =   1845
      End
      Begin VB.TextBox txt_USU_USUARIO 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   15
         TabIndex        =   4
         Top             =   240
         Width           =   1845
      End
      Begin VB.TextBox txtSenha 
         Alignment       =   1  'Right Justify
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   570
         Width           =   1845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Confirme senha:"
         Height          =   195
         Left            =   2970
         TabIndex        =   15
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nome :"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   600
         Width           =   510
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   60
      TabIndex        =   10
      Top             =   90
      Width           =   6525
      Begin VB.CommandButton btnPesquisaUsuario 
         BackColor       =   &H0080FF80&
         Caption         =   "..."
         Height          =   255
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   315
      End
      Begin VB.CommandButton CmdCancelaUsuario 
         Height          =   255
         Left            =   2250
         Picture         =   "frmGIFDisUsuario.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   270
         Width           =   315
      End
      Begin VB.CommandButton SGBotaoDBConfirmar 
         Height          =   255
         Left            =   1935
         Picture         =   "frmGIFDisUsuario.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   270
         Width           =   315
      End
      Begin VB.TextBox SGTcodigo 
         Height          =   315
         Left            =   780
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.CommandButton cmd_Acesso 
      Caption         =   "&Acesso"
      Enabled         =   0   'False
      Height          =   330
      Left            =   120
      TabIndex        =   9
      Top             =   2010
      Width           =   1275
   End
End
Attribute VB_Name = "frmGIFDisUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Public Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Public cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public cData_alteracao As String 'Data de alteracao vinda do registro a ser alterado ou excluido
Public Confirmada_Mudanca As Boolean 'Servirá para confirmar as mudanças de alteracoes dos campos na tela
Public sManutencao As Boolean ' Servirá para confirmar alteração dos campos
Public cTipo_Movimentacao As Integer  'Se for 1=Inclusão;2=Alteração

Private Sub btnPesquisaUsuario_Click()
Dim oTela As frmGIFDiscPesquisarUsuario
Set oTela = New frmGIFDiscPesquisarUsuario

    oTela.Show 1
    If oTela.ccodigo_pesquisa = "" Then
        Call Limpar_campos
        Me.SGTcodigo.Text = ""
        Me.SGTcodigo.SetFocus
    Else
        Me.SGTcodigo.Text = oTela.ccodigo_pesquisa
        SGBotaoDBConfirmar_Click
    End If
    Set oTela = Nothing

End Sub

Private Sub cmd_Acesso_Click()
Dim oTela As frmGIFDisUsuarioAcesso
Set oTela = New frmGIFDisUsuarioAcesso
    oTela.ccodigo_pesquisa = Me.SGTcodigo.Text
    oTela.txtNome.Text = Me.txt_USU_USUARIO.Text
    oTela.Show 1

    If oTela.ccodigo_pesquisa = "" Then
        Me.SGTcodigo.Text = ""
        Me.SGTcodigo.SetFocus
    Else
        Me.SGTcodigo.Text = oTela.ccodigo_pesquisa
        SGBotaoDBConfirmar_Click
    End If
    Unload oTela
    Set oTela = Nothing

End Sub

Private Sub CmdCancelaUsuario_Click()
    Me.Confirmada_Mudanca = False
    Me.Limpar_campos
    Me.Desabilitar_Campos
    Me.cmdnovo.Enabled = True
    Me.cmdsalvar.Enabled = False
    Me.cmdexcluir.Enabled = False
    Me.cmd_Acesso.Enabled = False
    Me.SGTcodigo.Locked = False
    Me.SGTcodigo.BackColor = &H80000005
    Me.SGTcodigo.SetFocus
    Me.cTipo_Movimentacao = 0
    Me.btnPesquisaUsuario.Enabled = True
    Me.SGBotaoDBConfirmar.Enabled = True
    Me.SGBotaoDBConfirmar.Default = False
End Sub

Private Sub cmdexcluir_Click()
Dim nRet As Integer
On Error GoTo Erro

nRet = MsgBox("Confirma exclusão?", vbQuestion & vbYesNo, Me.Caption)
'Se confirmou a exclusão:
If nRet = 6 Then
    Me.MousePointer = vbHourglass
    Call CCTemp.USUARIO_Excluir(sNomeBanco, Me.SGTcodigo.Text, cData_alteracao)
    Me.MousePointer = vbDefault
    Me.Confirmada_Mudanca = False
    Me.Limpar_campos
    Me.Desabilitar_Campos
    Me.cmdnovo.Enabled = True
    Me.cmdsalvar.Enabled = False
    Me.cmdexcluir.Enabled = False
    Me.SGTcodigo.Locked = False
    Me.SGTcodigo.BackColor = &H80000005
    Me.SGTcodigo.SetFocus
    Me.cTipo_Movimentacao = 0
    Me.btnPesquisaUsuario.Enabled = True
    Me.SGBotaoDBConfirmar.Enabled = True
End If

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmdfechar_Click()
Unload Me
End Sub

Private Sub cmdnovo_Click()
Call Limpar_campos
Call Habilitar_Campos
Me.SGTcodigo.BackColor = &H8000000F
Me.SGTcodigo.Locked = True
Me.SGTcodigo.Text = "NOVO"
Me.txt_USU_USUARIO.SetFocus
cTipo_Movimentacao = 1
Confirmada_Mudanca = True

End Sub

Private Sub cmdsalvar_Click()
Dim sSenhaCripta As String

Dim otemp As ADODB.Recordset
Dim cCodigo As String
Dim nx As Integer

Me.MousePointer = vbHourglass
On Error GoTo Erro

If TXTSENHA.Text <> Me.txtConfNovasenha.Text Then
   MsgBox "SENHA DIGITADA NÃO CONFERE COM A CONFIRMADA!!", , Me.Caption
   Me.txtConfNovasenha.SetFocus
   Me.MousePointer = vbDefault
   Exit Sub
End If

If cTipo_Movimentacao = 1 Then 'Inclusão da Turma
   sSenhaCripta = Cripta(Me.TXTSENHA.Text)
   Set otemp = CCTempneUsuario.USUARIO_Incluir(sNomeBanco, _
                                               Me.SGTcodigo.Text, _
                                               Me.txt_USU_USUARIO.Text, _
                                               sSenhaCripta, _
                                               sUsuario)
   
   Me.SGTcodigo.Text = otemp.Fields.Item(0)

Else ' Alteração da Turma
   sSenhaCripta = Cripta(Me.TXTSENHA.Text)
   Call CCTempneUsuario.USUARIO_Alterar(sNomeBanco, _
                                        Me.SGTcodigo.Text, _
                                        Me.txt_USU_USUARIO.Text, _
                                        sSenhaCripta, _
                                        sUsuario, _
                                        cData_alteracao)
End If
Me.MousePointer = vbDefault
cCodigo = Me.SGTcodigo.Text
Call Desabilitar_Campos
Confirmada_Mudanca = False
cTipo_Movimentacao = 0
Me.SGTcodigo.Text = cCodigo
Me.cmdnovo.Enabled = True
Me.cmdexcluir.Enabled = False
Me.cmdsalvar.Enabled = False
Me.SGTcodigo.BackColor = &H80000005
Me.SGTcodigo.Locked = False
Me.SGTcodigo.SetFocus
Me.btnPesquisaUsuario.Enabled = True
Me.SGBotaoDBConfirmar.Enabled = True
Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
If Flag_ativo = True Then
   Exit Sub
End If
Me.Top = 0
Me.Left = 0
Flag_ativo = True
Call Limpar_campos
Call Desabilitar_Campos
Me.SGTcodigo.BackColor = &H80000005
Me.SGTcodigo.Locked = False
Me.SGTcodigo.Enabled = True
Me.SGTcodigo.SetFocus
Me.cmdexcluir.Enabled = False
Me.cmdsalvar.Enabled = False
Confirmada_Mudanca = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' para funcionar , tem que mudar o keyPreviwe=true
If KeyCode = 13 Then
'   If Me.ActiveControl.TabIndex > 0 Then
      SendKeys "{TAB}"
'   End If
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

Private Sub SGBotaoDBConfirmar_Click()
On Error GoTo Erro
Me.SGBotaoDBConfirmar.Default = False
Me.MousePointer = vbHourglass

If Trim(Len(Me.SGTcodigo.Text)) = 0 Then
   MsgBox "Digite um valor para confirmar o código de Usuario", , Me.Caption
   Me.MousePointer = vbDefault
   Me.SGTcodigo.SetFocus
   Exit Sub
End If

SGTcodigo.Text = Format(SGTcodigo, "000")

Set cRec = CCTempneUsuario.USUARIO_Consultar(sNomeBanco, Me.SGTcodigo.Text)

If cRec Is Nothing Then
   If cTipo_Movimentacao = 0 Then
      MsgBox "Não Existe Usuario com este código, Tente Outro!"
   End If
   If cTipo_Movimentacao = 1 Then
      Habilitar_Campos
      Me.txt_USU_USUARIO.SetFocus
      Confirmada_Mudanca = True
      Me.cmdexcluir.Enabled = True
   End If
   Exit Sub
End If
If cRec.RecordCount > 0 Then
   Call Habilitar_Campos
   Call Carregar_campos
   Me.SGTcodigo.BackColor = &H8000000F
   Me.SGTcodigo.Locked = True
   Me.txt_USU_USUARIO.SetFocus
   Confirmada_Mudanca = True
   cTipo_Movimentacao = 2
   Me.cmdexcluir.Enabled = True
   Me.cmd_Acesso.Enabled = True
   Me.btnPesquisaUsuario.Enabled = False
   Me.SGBotaoDBConfirmar.Enabled = False
Else
   If cTipo_Movimentacao = 1 Then
      Habilitar_Campos
      Me.txt_USU_USUARIO.SetFocus
      Confirmada_Mudanca = True
   End If
   If cTipo_Movimentacao = 0 Then
      MsgBox "Não existe Usuario com este código, Tente Outro!"
   End If
End If
Me.MousePointer = vbDefault
Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub
Public Function CCTemp() As neUsuario
     Set CCTemp = New neUsuario
End Function
Function Limpar_campos()
    Me.SGTcodigo.Text = ""
    Me.txt_USU_USUARIO.Text = ""
    Me.TXTSENHA.Text = ""
    Me.txtConfNovasenha.Text = ""
End Function
Function Habilitar_Campos()
    Me.SGTcodigo.Locked = False
    Me.txt_USU_USUARIO.Locked = False
    Me.TXTSENHA.Locked = False
    Me.txtConfNovasenha.Locked = False
    
    Me.SGTcodigo.BackColor = &H80000005
    Me.txt_USU_USUARIO.BackColor = &H80000005
    Me.TXTSENHA.BackColor = &H80000005
    Me.txtConfNovasenha.BackColor = &H80000005

End Function
Function Desabilitar_Campos()
    Me.SGTcodigo.Locked = True
    Me.txt_USU_USUARIO.Locked = True
    Me.TXTSENHA.Locked = True
    Me.txtConfNovasenha.Locked = True

    Me.SGTcodigo.BackColor = &H8000000F
    Me.txt_USU_USUARIO.BackColor = &H8000000F
    Me.TXTSENHA.BackColor = &H8000000F
    Me.txtConfNovasenha.BackColor = &H8000000F

End Function
Function Carregar_campos()
Dim nx As Integer
Me.SGTcodigo.Text = cRec!USU_CODIGO
Me.txt_USU_USUARIO.Text = cRec!USU_USUARIO
Me.TXTSENHA.Text = IIf(IsNull(cRec!USU_SENHA), "", cRec!USU_SENHA)
cData_alteracao = IIf(IsNull(cRec!USU_DTA), "", cRec!USU_DTA)
End Function

Private Sub SGTcodigo_GotFocus()
SGBotaoDBConfirmar.Default = True
End Sub

Private Sub txt_USU_USUARIO_Change()
Call Confirma_Mudanca
End Sub

Private Sub txtConfNovasenha_Change()
Call Confirma_Mudanca
End Sub

Private Sub txtSenha_Change()
Call Confirma_Mudanca
End Sub

Private Function Confirma_Mudanca()

If Confirmada_Mudanca And Me.cmdsalvar.Visible = True Then
   Me.cmdexcluir.Enabled = False
   Me.cmdnovo.Enabled = False
   Me.cmdsalvar.Enabled = True
End If

End Function


