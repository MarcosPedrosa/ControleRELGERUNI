VERSION 5.00
Begin VB.Form frmGIFMntFormularios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manutenção dos Formulários do Sistema"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8190
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4740
      Width           =   1275
   End
   Begin VB.CommandButton cmdsalvar 
      Caption         =   "&Salvar"
      Height          =   330
      Left            =   5535
      TabIndex        =   9
      Top             =   4740
      Width           =   1275
   End
   Begin VB.CommandButton cmdexcluir 
      Caption         =   "&Excluir"
      Height          =   330
      Left            =   60
      TabIndex        =   14
      Top             =   4740
      Width           =   1275
   End
   Begin VB.CommandButton cmdnovo 
      Caption         =   "&Novo"
      Height          =   330
      Left            =   4230
      TabIndex        =   13
      Top             =   4740
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   8055
      Begin VB.CommandButton btnPesquisa 
         BackColor       =   &H00C0FFC0&
         Caption         =   "..."
         Height          =   255
         Left            =   1500
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton SGBotaoDBCancelar 
         Height          =   255
         Left            =   2160
         Picture         =   "frmGIFMntFormularios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   300
         Width           =   315
      End
      Begin VB.CommandButton SGBotaoDBConfirmar 
         Height          =   255
         Left            =   1830
         Picture         =   "frmGIFMntFormularios.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   315
      End
      Begin VB.TextBox TXT_FOR_CODIGO 
         Height          =   315
         Left            =   780
         MaxLength       =   6
         TabIndex        =   0
         Top             =   270
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.PictureBox vst_formulario 
      Height          =   3915
      Left            =   60
      ScaleHeight     =   3855
      ScaleWidth      =   7995
      TabIndex        =   15
      Top             =   780
      Width           =   8055
      Begin VB.PictureBox vsElastic2 
         Appearance      =   0  'Flat
         Height          =   3540
         Left            =   45
         ScaleHeight     =   3510
         ScaleWidth      =   7935
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   330
         Width           =   7965
         Begin ComctlLib.TreeView TreeView1 
            Height          =   3435
            Left            =   90
            TabIndex        =   17
            Top             =   60
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   6059
            _Version        =   327682
            LabelEdit       =   1
            Style           =   7
            Appearance      =   1
         End
      End
      Begin VB.PictureBox vsElastic1 
         Appearance      =   0  'Flat
         Height          =   3540
         Left            =   8700
         ScaleHeight     =   3510
         ScaleWidth      =   7935
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   330
         Width           =   7965
         Begin VB.TextBox TXT_FOR_NOME_EDITOR 
            Height          =   315
            Left            =   2850
            MaxLength       =   20
            TabIndex        =   5
            Top             =   1020
            Width           =   3765
         End
         Begin VB.TextBox TXT_FOR_NOME_FORM 
            Height          =   315
            Left            =   2850
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1440
            Width           =   3765
         End
         Begin VB.TextBox TXT_FOR_DESCRICAO 
            Height          =   315
            Left            =   2850
            MaxLength       =   50
            TabIndex        =   7
            Top             =   1860
            Width           =   3765
         End
         Begin VB.TextBox TXT_FOR_GRUPO 
            Height          =   315
            Left            =   2850
            MaxLength       =   10
            TabIndex        =   4
            Top             =   600
            Width           =   1425
         End
         Begin VB.TextBox TXT_FOR_ACESSO 
            Height          =   315
            Left            =   2850
            TabIndex        =   8
            Top             =   2310
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Ordem menu.:"
            Height          =   195
            Left            =   1170
            TabIndex        =   23
            Top             =   2340
            Width           =   990
         End
         Begin VB.Label Label43 
            Caption         =   "Descrição no Menu.:"
            Height          =   195
            Left            =   1170
            TabIndex        =   22
            Top             =   1920
            Width           =   1605
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nome Menu.:"
            Height          =   195
            Left            =   1170
            TabIndex        =   21
            Top             =   1080
            Width           =   960
         End
         Begin VB.Label Label4 
            Caption         =   "Nome do formulário.:"
            Height          =   195
            Left            =   1170
            TabIndex        =   20
            Top             =   1500
            Width           =   1545
         End
         Begin VB.Label Label5 
            Caption         =   "Grupo no Menu.:"
            Height          =   195
            Left            =   1170
            TabIndex        =   19
            Top             =   660
            Width           =   1605
         End
      End
   End
   Begin ComctlLib.ImageList Iml_sinais 
      Left            =   0
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGIFMntFormularios.frx":0294
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmGIFMntFormularios.frx":05AE
            Key             =   "vermelho"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGIFMntFormularios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Private Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Private cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public cData_alteracao As String 'Data de alteracao vinda do registro a ser alterado ou excluido
Public Confirma_Mudanca As Boolean 'Servirá para confirmar as mudanças de alteracoes dos campos na tela
Public sManutencao As Boolean ' Servirá para confirmar alteração dos campos
Public cTipo_Movimentacao As Integer  'Se for 1=Inclusão;2=Alteração
Private i As Integer, j As Integer
Private PaiTre As Node
Private Parent As Node
Private Parent3 As Node
Private Parent5 As Node
Private Parent7 As Node
Private Parent9 As Node
Private rs As ADODB.Recordset
Private sGrupo As String

Function Carregar_campos()
Dim nx As Integer
Dim sdata As String

Me.TXT_FOR_CODIGO.Text = IIf(IsNull(cRec!for_codigo), "", cRec!for_codigo)
Me.TXT_FOR_NOME_EDITOR.Text = IIf(IsNull(cRec!FOR_NOME_EDITOR), "", cRec!FOR_NOME_EDITOR)
Me.TXT_FOR_ACESSO.Text = IIf(IsNull(cRec!FOR_ACESSO), "", cRec!FOR_ACESSO)
Me.TXT_FOR_GRUPO.Text = IIf(IsNull(cRec!FOR_GRUPO), "", cRec!FOR_GRUPO)
Me.TXT_FOR_NOME_FORM.Text = IIf(IsNull(cRec!FOR_NOME_FORM), "", cRec!FOR_NOME_FORM)
Me.TXT_FOR_DESCRICAO.Text = IIf(IsNull(cRec!FOR_DESCRICAO), "", cRec!FOR_DESCRICAO)
cData_alteracao = IIf(IsNull(cRec!FOR_DTA), "", cRec!FOR_DTA)
End Function
Function Limpar_campos()
    Me.TXT_FOR_CODIGO.Text = ""
    Me.TXT_FOR_GRUPO.Text = ""
    Me.TXT_FOR_NOME_FORM.Text = ""
    Me.TXT_FOR_DESCRICAO.Text = ""
    Me.TXT_FOR_NOME_EDITOR.Text = ""
    Me.TXT_FOR_ACESSO.Text = ""
End Function
Function Habilitar_Campos()
     
    Me.TXT_FOR_CODIGO.Enabled = True
    Me.TXT_FOR_GRUPO.Enabled = True
    Me.TXT_FOR_NOME_FORM.Enabled = True
    Me.TXT_FOR_DESCRICAO.Enabled = True
    Me.TXT_FOR_NOME_EDITOR.Enabled = True
    Me.TXT_FOR_ACESSO.Enabled = True
        
    Me.TXT_FOR_CODIGO.BackColor = &H80000005
    Me.TXT_FOR_GRUPO.BackColor = &H80000005
    Me.TXT_FOR_NOME_FORM.BackColor = &H80000005
    Me.TXT_FOR_DESCRICAO.BackColor = &H80000005
    Me.TXT_FOR_NOME_EDITOR.BackColor = &H80000005
    Me.TXT_FOR_ACESSO.BackColor = &H80000005

End Function
Function Desabilitar_Campos()
    Me.TXT_FOR_CODIGO.Enabled = False
    Me.TXT_FOR_GRUPO.Enabled = False
    Me.TXT_FOR_NOME_FORM.Enabled = False
    Me.TXT_FOR_DESCRICAO.Enabled = False
    Me.TXT_FOR_NOME_EDITOR.Enabled = False
    Me.TXT_FOR_ACESSO.Enabled = False
    
    Me.TXT_FOR_CODIGO.BackColor = &H8000000F
    Me.TXT_FOR_GRUPO.BackColor = &H8000000F
    Me.TXT_FOR_NOME_FORM.BackColor = &H8000000F
    Me.TXT_FOR_DESCRICAO.BackColor = &H8000000F
    Me.TXT_FOR_NOME_EDITOR.BackColor = &H8000000F
    Me.TXT_FOR_ACESSO.BackColor = &H8000000F

End Function
Public Function Confirmar_Mudanca()
If Confirma_Mudanca And Me.cmdsalvar.Visible = True Then
   Me.cmdexcluir.Enabled = False
   Me.cmdnovo.Enabled = False
   Me.cmdsalvar.Enabled = True
End If

End Function

Private Sub btnPesquisa_Click()
Dim oTela As frmGIFMntPesquisarFormulario
Set oTela = New frmGIFMntPesquisarFormulario

    oTela.Show 1
    If oTela.ccodigo_pesquisa = "" Then
        Me.TXT_FOR_CODIGO.Text = ""
        Me.TXT_FOR_CODIGO.SetFocus
    Else
        Me.TXT_FOR_CODIGO.Text = oTela.ccodigo_pesquisa
        SGBotaoDBConfirmar_Click
    End If
    Unload oTela: Set oTela = Nothing
End Sub


Private Sub Carrega_Treview()

On Error GoTo Erro
Dim XCS As Double
Set rs = New ADODB.Recordset

Me.MousePointer = vbHourglass

Set rs = CCTempneFormulario.TipoForm_Consultar(sNomeBanco)

If Not rs.BOF Then
   Me.TreeView1.Nodes.Clear
   rs.MoveFirst
   sGrupo = Mid$(rs!FOR_GRUPO, 1, 2)
   Set PaiTre = TreeView1.Nodes.Add(, , "A0000000000", "Menu do sistema")
   PaiTre.Expanded = True
'   Set Parent = TreeView1.Nodes.Add(, , "j" + rs!FOR_GRUPO, Mid$(rs!FOR_GRUPO, 1, 2) & "-" & rs!FOR_NOME_EDITOR, IIf(Mid$(rs!FOR_GRUPO, 3, 2) = "00", 1, 2), IIf(Mid$(rs!FOR_GRUPO, 3, 2), 1, 2))
   While Not rs.EOF
        i = 0: j = 0
        
        sGrupo = Mid$(rs!FOR_GRUPO, 1, 2)
        
        If Mid$(rs!FOR_GRUPO, 9, 2) <> "00" Then
           While Not rs.EOF
             Set Parent9 = TreeView1.Nodes.Add(Parent7, tvwChild, _
                                             Format(rs!for_codigo, "000") & "K" + rs!FOR_GRUPO, _
                                             "[" & Format(rs!FOR_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]")
             rs.MoveNext
           Wend
        End If
        
        If Not rs.EOF Then
           If Mid$(rs!FOR_GRUPO, 7, 2) <> "00" Then
              While Not rs.EOF
                Set Parent7 = TreeView1.Nodes.Add(Parent5, tvwChild, _
                                                  Format(rs!for_codigo, "000") & "K" + rs!FOR_GRUPO, _
                                                  "[" & Format(rs!FOR_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]")
                rs.MoveNext
             Wend
           End If
        End If
        
        If Not rs.EOF Then
           If Mid$(rs!FOR_GRUPO, 5, 2) <> "00" Then
              While Not rs.EOF
                Set Parent5 = TreeView1.Nodes.Add(Parent3, tvwChild, _
                                                  Format(rs!for_codigo, "000") & "K" + rs!FOR_GRUPO, _
                                                  "[" & Format(rs!FOR_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]")
                rs.MoveNext
              Wend
           End If
        End If
        
        If Not rs.EOF Then
           If Mid$(rs!FOR_GRUPO, 3, 2) <> "00" Then
              While Not rs.EOF And Mid$(rs!FOR_GRUPO, 3, 2) <> "00"
                Set Parent3 = TreeView1.Nodes.Add(Parent, tvwChild, _
                                                  Format(rs!for_codigo, "000") & "K" + rs!FOR_GRUPO, _
                                                  "[" & Format(rs!FOR_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]")
                rs.MoveNext
                If rs.EOF Then GoTo SAIDA_LOOP
                If Mid$(rs!FOR_GRUPO, 5, 2) <> "00" Then
                   While Not rs.EOF And Mid$(rs!FOR_GRUPO, 5, 2) <> "00"
                     Set Parent5 = TreeView1.Nodes.Add(Parent3, tvwChild, _
                                                       Format(rs!for_codigo, "000") & "K" + rs!FOR_GRUPO, _
                                                       "[" & Format(rs!FOR_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]")
                     rs.MoveNext
                     If rs.EOF Then GoTo SAIDA_LOOP
                     If Mid$(rs!FOR_GRUPO, 7, 2) <> "00" Then
                        While Not rs.EOF And Mid$(rs!FOR_GRUPO, 7, 2) <> "00"
                          Set Parent7 = TreeView1.Nodes.Add(Parent5, tvwChild, _
                                                            Format(rs!for_codigo, "000") & "K" + rs!FOR_GRUPO, _
                                                            "[" & Format(rs!FOR_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]")
                          rs.MoveNext
                          If rs.EOF Then GoTo SAIDA_LOOP
                          If Mid$(rs!FOR_GRUPO, 9, 2) <> "00" Then
                             While Not rs.EOF And Mid$(rs!FOR_GRUPO, 9, 2) <> "00"
                               Set Parent9 = TreeView1.Nodes.Add(Parent7, tvwChild, _
                                                                 Format(rs!for_codigo, "000") & "K" + rs!FOR_GRUPO, _
                                                                 "[" & Format(rs!FOR_ACESSO, "00") & " - " & rs!FOR_NOME_EDITOR & " - " & rs!FOR_DESCRICAO & "]")
                               rs.MoveNext
                             Wend
                          End If
                       Wend
                     End If
                   Wend
                End If
              Wend
           End If
        End If
SAIDA_LOOP:
        If Not rs.EOF Then
           Set Parent = TreeView1.Nodes.Add(PaiTre, tvwChild, Format(rs!for_codigo, "000") & "J" + rs!FOR_GRUPO, Mid$(rs!FOR_GRUPO, 1, 2) & "-" & rs!FOR_DESCRICAO)
'           Parent.Expanded = True
           rs.MoveNext
        End If
    Wend
End If



Me.MousePointer = vbDefault
Set rs = Nothing

Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub


Private Sub TreeView1_DblClick()
'Me.TXT_FOR_GRUPO.Text = Mid$(Me.TreeView1.SelectedItem.Key, 2, Len(Me.TreeView1.SelectedItem.Key))
'Me.TXT_FOR_NOME_EDITOR.Text = Mid$(Me.TreeView1.SelectedItem.Key, 2, Len(Me.TreeView1.SelectedItem.Key))
'Me.TXT_FOR_NOME_FORM.Text = Mid$(Me.TreeView1.SelectedItem.Key, 2, Len(Me.TreeView1.SelectedItem.Key))
'Me.TXT_FOR_DESCRICAO.Text = ""
'Me.TXT_FOR_ACESSO.Text = ""
If Mid$(Me.TreeView1.SelectedItem.Key, 1, 1) = "A" Then Exit Sub
Me.TXT_FOR_CODIGO.Text = Mid$(Me.TreeView1.SelectedItem.Key, 1, 3)
SGBotaoDBConfirmar_Click
'MsgBox Me.TreeView1.Nodes.Item(1).Key
End Sub
Private Sub TXT_FOR_DESCRICAO_Click()
Confirmar_Mudanca
End Sub

Private Sub TXT_FOR_GRUPO_Change()
Confirmar_Mudanca
End Sub

Private Sub TXT_FOR_NOME_FORM_Change()
Confirmar_Mudanca
End Sub

Private Sub cmdexcluir_Click()
Dim nRet As Integer
On Error GoTo Erro

nRet = MsgBox("Confirma exclusão?", vbQuestion & vbYesNo, Me.Caption)
'Se confirmou a exclusão:
If nRet = 6 Then
    Me.MousePointer = vbDefault
    Call CCTempneFormulario.TipoForm_Excluir(sNomeBanco, _
                                             Me.TXT_FOR_CODIGO.Text, _
                                             sUsuario, _
                                             cData_alteracao)
    Me.MousePointer = vbDefault
    Me.Confirma_Mudanca = False
    Me.Limpar_campos
    Me.Desabilitar_Campos
    Me.cmdnovo.Enabled = True
    Me.cmdsalvar.Enabled = False
    Me.cmdexcluir.Enabled = False
    Me.TXT_FOR_CODIGO.Enabled = True
    Me.TXT_FOR_CODIGO.BackColor = &H80000005
    Me.TXT_FOR_CODIGO.SetFocus
    Me.cTipo_Movimentacao = 0
    Me.btnPesquisa.Enabled = True
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
Me.vst_formulario.CurrTab = 1
Call Habilitar_Campos
Me.TXT_FOR_CODIGO.BackColor = &H8000000F
Me.TXT_FOR_CODIGO.Enabled = False
Me.TXT_FOR_CODIGO.Text = "NOVO"
cTipo_Movimentacao = 1
Confirma_Mudanca = True
Me.TXT_FOR_GRUPO.SetFocus
End Sub

Private Sub cmdsalvar_Click()

Dim otemp As ADODB.Recordset
Dim cCodigo As String
Dim nx As Integer
Dim cenderecoCom As String
Dim cBairroCom As String
Dim cMunicipioCom As String
Dim ccepCom As String
Dim cufCom As String
Dim ctelefoneCom As String
Dim cfaxCom As String
Dim sdata As String

Me.MousePointer = vbHourglass
On Error GoTo Erro

If cTipo_Movimentacao = 1 Then 'Inclusão do FORMULARIO
   Set otemp = CCTempneFormulario.TipoForm_Incluir(sNomeBanco, _
                                                   Me.TXT_FOR_CODIGO.Text, _
                                                   Me.TXT_FOR_NOME_EDITOR.Text, _
                                                   Me.TXT_FOR_NOME_FORM.Text, _
                                                   Me.TXT_FOR_DESCRICAO.Text, _
                                                   Me.TXT_FOR_GRUPO.Text, _
                                                   Me.TXT_FOR_ACESSO.Text, _
                                                   sUsuario)
   Me.TXT_FOR_CODIGO.Text = otemp.Fields.Item(0)
   Set otemp = Nothing
Else ' Alteração do FORMULARIO
   Call CCTempneFormulario.TipoForm_Alterar(sNomeBanco, _
                                            Me.TXT_FOR_CODIGO.Text, _
                                            Me.TXT_FOR_NOME_EDITOR.Text, _
                                            Me.TXT_FOR_NOME_FORM.Text, _
                                            Me.TXT_FOR_DESCRICAO.Text, _
                                            Me.TXT_FOR_GRUPO.Text, _
                                            Me.TXT_FOR_ACESSO.Text, _
                                            sUsuario, _
                                            cData_alteracao)
End If
Me.MousePointer = vbDefault
cCodigo = Me.TXT_FOR_CODIGO.Text
Desabilitar_Campos
Confirma_Mudanca = False
cTipo_Movimentacao = 0
Me.TXT_FOR_CODIGO.Text = cCodigo
Me.cmdnovo.Enabled = True
Me.cmdexcluir.Enabled = False
Me.cmdsalvar.Enabled = False
Me.TXT_FOR_CODIGO.BackColor = &H80000005
Me.TXT_FOR_CODIGO.Enabled = True
Me.TXT_FOR_CODIGO.SetFocus
Me.btnPesquisa.Enabled = True
Me.SGBotaoDBConfirmar.Enabled = True

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub TXT_FOR_DESCRICAO_Change()
Confirmar_Mudanca
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
Call Carrega_Treview
Me.TXT_FOR_CODIGO.BackColor = &H80000005
Me.TXT_FOR_CODIGO.Enabled = True
Me.TXT_FOR_CODIGO.SetFocus
Me.cmdexcluir.Enabled = False
Me.cmdsalvar.Enabled = False
Confirma_Mudanca = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' para funcionar , tem que mudar o keyPreviwe=true
If KeyCode = 13 Then
   If Me.ActiveControl.TabIndex > 3 Then
      SendKeys "{TAB}"
   End If
ElseIf KeyCode = 27 Then
'   If Me.ActiveControl.TabIndex < 8 Then
'      If Me.cmdsalvar.Enabled = True Then
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
Me.vst_formulario.CurrTab = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cmdsalvar.Enabled = True Then
    If 7 = MsgBox("Possiveis alterações foram realizadas sem salvar,Deseja abandonar?", 32 + 4) Then
        'respondeu não
        Cancel = True
    End If
End If

End Sub

Private Sub SGBotaoDBCancelar_Click()
Me.SGBotaoDBConfirmar.Default = False
Call Limpar_campos
Call Desabilitar_Campos
Me.TXT_FOR_CODIGO.BackColor = &H80000005
Me.TXT_FOR_CODIGO.Enabled = True
Me.TXT_FOR_CODIGO.SetFocus
Confirma_Mudanca = False
cTipo_Movimentacao = 0
Me.cmdexcluir.Enabled = False
Me.cmdnovo.Enabled = True
Me.cmdsalvar.Enabled = False
Me.btnPesquisa.Enabled = True
Me.SGBotaoDBConfirmar.Enabled = True
Me.TXT_FOR_CODIGO.SetFocus
End Sub

Private Sub SGBotaoDBConfirmar_Click()
On Error GoTo Erro
Me.SGBotaoDBConfirmar.Default = False
Me.MousePointer = vbHourglass

If Trim(Len(Me.TXT_FOR_CODIGO.Text)) = 0 Then
   MsgBox "Digite um valor para confirmar o código do formulário", , Me.Caption
   Me.MousePointer = vbDefault
   Me.TXT_FOR_CODIGO.SetFocus
   Exit Sub
End If

Me.vst_formulario.CurrTab = 1

TXT_FOR_CODIGO.Text = Format(TXT_FOR_CODIGO, "000")

Set cRec = CCTempneFormulario.TipoForm_Consultar(sNomeBanco, Me.TXT_FOR_CODIGO.Text)

If cRec Is Nothing Then
   If cTipo_Movimentacao = 0 Then
      MsgBox "Não Existe Formulário com este Código, Tente Outro!"
   End If
   If cTipo_Movimentacao = 1 Then
      Call Habilitar_Campos
      Me.TXT_FOR_GRUPO.SetFocus
      Confirma_Mudanca = True
      Me.cmdexcluir.Enabled = True
   End If
   Set cRec = Nothing
   Exit Sub
End If
If cRec.RecordCount > 0 Then
   Call Habilitar_Campos
   Call Carregar_campos
   Me.TXT_FOR_CODIGO.BackColor = &H8000000F
   Me.TXT_FOR_CODIGO.Enabled = False
   Me.TXT_FOR_NOME_EDITOR.SetFocus
   Confirma_Mudanca = True
   cTipo_Movimentacao = 2
   Me.cmdexcluir.Enabled = True
   Me.btnPesquisa.Enabled = False
   Me.SGBotaoDBConfirmar.Enabled = False
Else
   If cTipo_Movimentacao = 1 Then
      Habilitar_Campos
      Me.TXT_FOR_GRUPO.SetFocus
      Confirma_Mudanca = True
   End If
   If cTipo_Movimentacao = 0 Then
      MsgBox "Não Existe TIPO OCORRENCIA com este Código, Tente Outro!"
   End If
End If
Set cRec = Nothing
Me.MousePointer = vbDefault
Exit Sub

Erro:
Set cRec = Nothing
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub

Private Sub TXT_FOR_CODIGO_GotFocus()
    SGBotaoDBConfirmar.Default = True
End Sub

Private Sub TXT_FOR_CODIGO_LostFocus()
    SGBotaoDBConfirmar.Default = False
End Sub

Private Sub TXT_FOR_ACESSO_Change()
Confirmar_Mudanca
End Sub

Private Sub TXT_FOR_NOME_EDITOR_Change()
Confirmar_Mudanca
End Sub

'Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
'     MsgBox "Colapsing node: " & Node.Text
'End Sub

'Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
'     MsgBox "Expanding node: " & Node.Text
'End Sub

'Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
'     If Node.Checked Then
'          MsgBox "Node " & Node.Text & " was checked"
'     Else
'          MsgBox "Node " & Node.Text & " was Unchecked"
'     End If
'End Sub


Private Sub vst_formulario_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
If NewTab = 0 Then
   SGBotaoDBCancelar_Click
   Call Carrega_Treview
End If
End Sub


