VERSION 5.00
Begin VB.Form frmGIFIUniDigitacaoDescontoExtra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digitação de Descontos Extras"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7185
   Begin VB.Frame Frame2 
      Caption         =   "Selecione o funcionário e o valor a descontar"
      Height          =   1575
      Left            =   60
      TabIndex        =   11
      Top             =   1680
      Width           =   7035
      Begin VB.CommandButton cmd_Limpar_Cod_Func 
         BackColor       =   &H000000FF&
         Caption         =   "X"
         Height          =   255
         Left            =   2250
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar codigo do funcionário"
         Top             =   390
         Width           =   285
      End
      Begin VB.CheckBox chk_MFU_TIPO 
         Caption         =   "Desconto p/folha"
         Height          =   255
         Left            =   1140
         TabIndex        =   8
         Top             =   1170
         Value           =   1  'Checked
         Width           =   1545
      End
      Begin VB.TextBox TXT_SALDO_ATUAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0,00"
         Top             =   750
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox TXT_SALDO_ANT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   3390
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0,00"
         Top             =   750
         Width           =   1155
      End
      Begin VB.TextBox TXT_DESCONTO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Text            =   "0,00"
         Top             =   750
         Width           =   1185
      End
      Begin VB.TextBox txtFuncionario 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2550
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   360
         Width           =   4335
      End
      Begin VB.CommandButton cmd_pesquisa 
         Caption         =   "..."
         Height          =   255
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   390
         Width           =   285
      End
      Begin VB.TextBox txtCODFUNC 
         Height          =   315
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   4
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Atu."
         Height          =   195
         Left            =   4830
         TabIndex        =   18
         Top             =   810
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Ant."
         Height          =   195
         Left            =   2535
         TabIndex        =   16
         Top             =   810
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Desconto.:"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   810
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Funcionário :"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   390
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3300
      Width           =   1275
   End
   Begin VB.CommandButton cmd_confirmar 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Confirmar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4470
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3300
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecione a empresa/Pesquisar Mov digitado"
      Height          =   1515
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7035
      Begin VB.CommandButton cmd_confirmar_NF 
         BackColor       =   &H00FF8080&
         Caption         =   "Consultar Digitação"
         Height          =   465
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "pesquisar movimentos digitados no Periodo/Coligada"
         Top             =   870
         Width           =   1605
      End
      Begin VB.ComboBox cbo_mes_ano 
         BackColor       =   &H00C0FFFF&
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
         Height          =   555
         Left            =   150
         TabIndex        =   2
         Text            =   "01/2010"
         Top             =   780
         Width           =   1935
      End
      Begin VB.ComboBox cbo_coligada 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   6765
      End
   End
End
Attribute VB_Name = "frmGIFIUniDigitacaoDescontoExtra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Private sAnoMes As String 'contem o ano mes referente ao arquivo da coligada
Private sPercentual As String 'contem o percentual referente ao arquivo da coligada
Private sMovSituacao As String 'cotem a situacao do arquivo da coligada 0=aberto, 1=fechado
Private cRec As ADODB.Recordset
Private cColigada As String ' contera a coligada escolhida com o movimento do grid preenchido
Private Sub cbo_coligada_Change()
Call Confirmar_Dados_coligada
End Sub

Private Sub cbo_coligada_Click()
Call Confirmar_Dados_coligada
End Sub

Private Sub cbo_mes_ano_Change()
sAnoMes = Me.cbo_mes_ano.ItemData(Me.cbo_mes_ano.ListIndex)
End Sub

Private Sub cbo_mes_ano_Click()
sAnoMes = Me.cbo_mes_ano.ItemData(Me.cbo_mes_ano.ListIndex)
End Sub
Private Sub cmd_Confirmar_Click()
Dim nx As Integer
Dim rs As ADODB.Recordset
Dim RESPOSTA As Integer

On Error GoTo Erro


If VBA.CDbl(Me.TXT_DESCONTO.Text) > VBA.CDbl(Me.TXT_SALDO_ANT.Text) Then
   MsgBox "Valor a ser descontado maior que o Saldo, Redigite!"
   Me.TXT_DESCONTO.SetFocus
   Exit Sub
End If

RESPOSTA = MsgBox("Confirma desconto deste funcionário ?", 20, "Sim/Não?")

If RESPOSTA = 7 Then Exit Sub

Rem CRITICA PARA SABER SE O VALOR DIGITADO FOR MAIOR QUE O SALDO

If VBA.CDbl(Me.TXT_DESCONTO.Text) > VBA.CDbl(Me.TXT_SALDO_ANT.Text) Then
   MsgBox "Valor a ser descontado maior que o saldo devedor do funcionário. Redigite!"
   Me.TXT_DESCONTO.SetFocus
   Exit Sub
End If

Me.MousePointer = vbHourglass

Set rs = New ADODB.Recordset

rs.Fields.Append "CHAPA", ADODB.DataTypeEnum.adVarChar, 7
rs.Fields.Append "FUNCIONARIO", ADODB.DataTypeEnum.adVarChar, 30
rs.Fields.Append "DT_EVENTO", ADODB.DataTypeEnum.adVarChar, 10
rs.Fields.Append "HORA", ADODB.DataTypeEnum.adVarChar, 5
rs.Fields.Append "REFERENCIA", ADODB.DataTypeEnum.adDouble
rs.Fields.Append "VALOR", ADODB.DataTypeEnum.adDouble
rs.Fields.Append "TIPO", ADODB.DataTypeEnum.adInteger
rs.Fields.Append "SALDO_ANT", ADODB.DataTypeEnum.adDouble

rs.Open
rs.AddNew "CHAPA", IIf(IsNull(txtCODFUNC.Text), " ", Mid$(txtCODFUNC.Text, 1, 7))
rs.Fields("FUNCIONARIO").Value = IIf(IsNull(txtFuncionario.Text), " ", Mid$(txtFuncionario.Text, 1, 30))
rs.Fields("DT_EVENTO").Value = Format(Now(), "dd/mm/yyyy")
rs.Fields("HORA").Value = Format(Now(), "hh:mm")
rs.Fields("REFERENCIA").Value = "1.00"
rs.Fields("VALOR").Value = IIf(IsNull(TXT_DESCONTO.Text), " ", CDbl(Trim(TXT_DESCONTO.Text)))
If Me.chk_MFU_TIPO.Value = 1 Then
   rs.Fields("TIPO").Value = 1
Else
   rs.Fields("TIPO").Value = 2
End If
rs.Fields("SALDO_ANT").Value = IIf(IsNull(Me.TXT_SALDO_ANT.Text), " ", CDbl(Trim(TXT_SALDO_ANT.Text)))
rs.Update
       
Call CCTempneUniMvFun.MovFuncionario_Incluir(Mid$(Me.cbo_mes_ano.List(Me.cbo_mes_ano.ListIndex), 4, 4) & Mid$(Me.cbo_mes_ano.List(Me.cbo_mes_ano.ListIndex), 1, 2), _
                                             sPercentual, _
                                             Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex), _
                                             rs)
       
Me.MousePointer = vbDefault

MsgBox "Inclusão realizada com sucesso!"
Call Limpar_campos

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmd_confirmar_NF_Click()

Dim oTela As frmGIFIUniPesquisarDescontodigitado
Set oTela = New frmGIFIUniPesquisarDescontodigitado

Call Limpar_campos
Me.txtCODFUNC.Text = ""
oTela.cColigada = Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex)
oTela.cMesAno = Mid$(Me.cbo_mes_ano.List(Me.cbo_coligada.ListIndex), 4, 4) & _
                Mid$(Me.cbo_mes_ano.List(Me.cbo_coligada.ListIndex), 1, 2)
oTela.Show 1
If oTela.ccodigo_pesquisa = "" Or oTela.ccodigo_pesquisa = "CHAPA" Then
    Call Limpar_campos
    Me.txtCODFUNC.Text = ""
    Me.txtCODFUNC.SetFocus
Else
    Me.txtCODFUNC.Text = Format(oTela.ccodigo_pesquisa, "000000")
    Me.txtFuncionario.Text = oTela.cnome
    Me.TXT_DESCONTO.Text = oTela.nValordesc
    Me.TXT_SALDO_ANT.Text = oTela.nValorSaldo
    Me.cmd_confirmar.Enabled = True
    Me.TXT_DESCONTO.SetFocus
End If
Set oTela = Nothing

End Sub

Private Sub cmd_Limpar_Cod_Func_Click()
Me.txtCODFUNC.Text = ""
Me.txtFuncionario.Text = ""
Me.txtCODFUNC.Locked = False
Me.TXT_DESCONTO.Text = "0.00"
Me.TXT_SALDO_ANT.Text = "0.00"
Me.txtCODFUNC.SetFocus
End Sub

Private Sub cmd_pesquisa_Click()
Dim oTela As frmGIFIUniPesquisarFuncionario

Set oTela = New frmGIFIUniPesquisarFuncionario
    oTela.cColigada = Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex)
    oTela.Show 1
    If oTela.ccodigo_pesquisa = "" Then
        Me.txtCODFUNC.Text = ""
        Me.txtFuncionario.Text = ""
        Me.txtCODFUNC.SetFocus
    Else
        Me.txtCODFUNC.Text = Format(oTela.ccodigo_pesquisa, "000000")
        Me.txtFuncionario.Text = oTela.cnome
        Call Pesquisar_Saldo
        Me.TXT_DESCONTO.Text = ""
        Me.TXT_DESCONTO.SetFocus
    End If
    Set oTela = Nothing

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

Private Sub Form_Load()
Dim nx As Integer

Me.Top = 0
Me.Left = 0

Call carregar_coligada
Call carregar_Meses_Fechados

End Sub

Private Sub carregar_coligada()
Dim nx As Integer
Dim cRec As ADODB.Recordset

On Error GoTo Erro

Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass
Set cRec = rRec_cliente
Set cRec = CCTempneUniColigada.Coligada_Consultar(sBancoUnimed)

Me.cbo_coligada.Clear

nx = 0

If cRec.RecordCount > 0 Then
   cRec.MoveFirst
   While Not cRec.EOF
       If Not IsNull(cRec!TCO_CODIGO) Then
          Me.cbo_coligada.AddItem cRec!TCO_CODIGO & " - " & Trim(cRec!TCO_DESCRICAO)
          Me.cbo_coligada.ItemData(nx) = cRec!TCO_CODIGO
          nx = nx + 1
       End If
       cRec.MoveNext
   Wend
   If nx < 2 Then
      Me.cbo_coligada.ListIndex = 0
      Call Confirmar_Dados_coligada
   End If
Else
   MsgBox "Não existem empresas Coligadas, procure o responsável."
End If

Me.MousePointer = vbDefault

Set cRec = Nothing
Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub

Private Sub carregar_Meses_Fechados()
Dim nx As Integer
Dim cRec As ADODB.Recordset

On Error GoTo Erro

Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass

Set cRec = CCTempneUniMvFun.MovFuncionario_ConsMesFechado(Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex))

Me.cbo_mes_ano.Clear

nx = 0

If cRec.RecordCount > 0 Then
   cRec.MoveFirst
   While Not cRec.EOF
       If Not IsNull(cRec!MFU_ANO_MES) Then
          Me.cbo_mes_ano.AddItem Mid$(cRec!MFU_ANO_MES, 5, 2) & "/" & Mid$(cRec!MFU_ANO_MES, 1, 4)
          Me.cbo_mes_ano.ItemData(nx) = cRec!MFU_ANO_MES
          nx = nx + 1
       End If
       cRec.MoveNext
   Wend
   Me.cbo_mes_ano.ListIndex = 0
   If nx > 1 Then
      Me.cbo_mes_ano.Enabled = True
   End If
Else
   Me.cbo_mes_ano.AddItem "000000"
   Me.cbo_mes_ano.ListIndex = 0
   MsgBox "Não existem Fechamentos das empresas Coligadas, procure o responsável."
End If

Me.MousePointer = vbDefault

Set cRec = Nothing
Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub
Private Sub Confirmar_Dados_coligada()
Dim cRec As ADODB.Recordset

On Error GoTo Erro

Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass
Set cRec = rRec_cliente
Set cRec = CCTempneUniColigada.Coligada_Consultar(sBancoUnimed, Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex))
'lbl_Msg_Fechamento.Caption = ""

If cRec.RecordCount > 0 Then
   cRec.MoveFirst
   While Not cRec.EOF
       If cRec!TCO_CODIGO = Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex) Then
          sPercentual = cRec!TCO_DESCONTO
       End If
       cRec.MoveNext
   Wend
Else
   MsgBox "Não existem empresas Coligadas, procure o responsável."
End If

Me.MousePointer = vbDefault

Set cRec = Nothing
Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub


Private Sub TXT_DESCONTO_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
            Case vbKeyDelete
            Case vbKeyBack
            Case 97 To 99
            Case 65 To 67
            Case 42
            Case 61
                 KeyAscii = 0
                 Me.TXT_DESCONTO.Text = Me.TXT_SALDO_ANT.Text
            Case Else
                 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Select

'If KeyAscii = 13 Then
''    txtQtdValeEmpresa.SetFocus
'End If

Call Habilitar_confirma

End Sub
'
'Private Function Acesso_Funcionario()
'
'On Error GoTo erro
'
'Me.MousePointer = vbHourglass
'Set cRec = New ADODB.Recordset
'
'Set cRec = CCTempneTabRegPagto.TabRegPagto_Consultar_Funcionario(sBancoRM, _
'                                                                 Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex), _
'                                                                 Me.txtCODFUNC.Text)
'
'If cRec.RecordCount > 0 Then
'End If
'
'Me.MousePointer = vbDefault
'
'Exit Function
'
'erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
'
'End Function

Private Function Limpar_campos()

Me.txtCODFUNC.Text = ""
Me.txtFuncionario.Text = ""
Me.TXT_DESCONTO.Text = "0,00"
Me.TXT_SALDO_ANT.Text = "0,00"
'Me.TXT_SALDO_ATUAL.Text = "0,00"

End Function

Private Sub TXT_DESCONTO_LostFocus()
Me.TXT_DESCONTO.Text = Format(Me.TXT_DESCONTO.Text, "0.00")
End Sub

Private Sub txtCODFUNC_Change()
Call Habilitar_confirma
End Sub

Private Sub Habilitar_confirma()
If Val(Me.TXT_DESCONTO.Text) > 0 And _
   Len(txtCODFUNC.Text) > 0 Then
   If VBA.CDbl(Me.TXT_DESCONTO.Text) <= VBA.CDbl(Me.TXT_SALDO_ANT.Text) Then
      Me.cmd_confirmar.Enabled = True
   End If
Else
   Me.cmd_confirmar.Enabled = False
End If
End Sub

Function Pesquisar_Saldo()
On Error GoTo Erro

Me.MousePointer = vbHourglass
Set cRec = New ADODB.Recordset

Set cRec = CCTempneUniMvFun.MovFuncionario_Consulta_Saldo(Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex), _
                                                          Me.txtCODFUNC.Text)

If cRec.RecordCount > 0 Then
   Me.TXT_SALDO_ANT.Text = Format(cRec!SAL_SALDO, "0.00")
End If

Me.MousePointer = vbDefault

Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Function

Private Sub txtCODFUNC_KeyPress(KeyAscii As Integer)

Dim rs1 As ADODB.Recordset

On Error GoTo Erro

If KeyAscii = 13 Then

   Me.MousePointer = vbHourglass
   Set rs1 = New ADODB.Recordset
   
   Rem VERIFICAR OS FUNCIONARIOS QUE EXISTEM SALDO NO CADASTRO
   
   Set rs1 = CCTempneUniMvFun.RMFuncionario_Consulta("1", Trim(Me.txtCODFUNC.Text))
   
   If rs1.RecordCount > 0 Then
      Me.txtFuncionario.Text = rs1!NOME
      Me.txtCODFUNC.Locked = True
   Else
      MsgBox "Funcionário não encontrado, redigite e tente novamente ou pesquise pelo nome no botão ao lado!"
   End If
   
   Me.MousePointer = vbDefault
   Set rs1 = Nothing
   
End If

Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault
   
End Sub

