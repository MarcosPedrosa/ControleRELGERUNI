VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGIFIUniPesquisarDescontodigitado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelar Funcionários digitados manualmente"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Cancelamento da movimentação"
      Height          =   975
      Left            =   60
      TabIndex        =   10
      Top             =   4260
      Width           =   9285
      Begin VB.CommandButton cmd_cancelamento 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   555
         Left            =   8190
         Picture         =   "frmGIFIUniPesquisarDescontodigitado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cancelar Nota Fiscal"
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txtNome 
         BackColor       =   &H0080FFFF&
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
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   7065
      End
      Begin VB.Label Label2 
         Caption         =   "Nome :"
         Height          =   225
         Left            =   90
         TabIndex        =   12
         Top             =   390
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   10605
      Begin MSFlexGridLib.MSFlexGrid mfl_grid 
         Height          =   3465
         Left            =   90
         TabIndex        =   9
         Top             =   180
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   6112
         _Version        =   393216
         Cols            =   9
         AllowBigSelection=   0   'False
         TextStyle       =   3
         TextStyleFixed  =   2
         HighLight       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Lbl_Processo 
         BackColor       =   &H00C0FFFF&
         Caption         =   "AGUARDE PROCESSAMENTO..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1245
         Left            =   1980
         TabIndex        =   6
         Top             =   1410
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1275
   End
   Begin VB.CommandButton cmdSelecionar 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Selecionar"
      Height          =   330
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5220
      Width           =   1275
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   2
      Top             =   3870
      Width           =   1005
   End
   Begin VB.OptionButton Opt_nome 
      Caption         =   "Nome"
      Height          =   255
      Left            =   8370
      TabIndex        =   1
      Top             =   5370
      Value           =   -1  'True
      Width           =   795
   End
   Begin VB.OptionButton Opt_secao 
      Caption         =   "Seção"
      Height          =   255
      Left            =   9270
      TabIndex        =   0
      Top             =   5370
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Total registros : "
      Height          =   285
      Left            =   60
      TabIndex        =   8
      Top             =   3900
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Classificação:"
      Height          =   195
      Left            =   7230
      TabIndex        =   7
      Top             =   5400
      Width           =   975
   End
End
Attribute VB_Name = "frmGIFIUniPesquisarDescontodigitado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ccodigo_pesquisa As String 'Codigo escolhido pelo usuário.
Public cnome As String 'Nome do escolhido pelo usuário.
Public rs As ADODB.Recordset
Public nTeclou_Enter As Integer
Public cColigada As String
Public nValordesc As Double
Public nValorSaldo As Double
Public cMesAno As String

Private Sub cmd_cancelamento_Click()
Dim nx As Integer
Dim rs As ADODB.Recordset
Dim RESPOSTA As Integer

On Error GoTo Erro

If Len(Trim(Me.txtNome.Text)) = 0 Then Exit Sub

RESPOSTA = MsgBox("Confirma cancelamento deste funcionário ?", 20, "Sim/Não?")

If RESPOSTA = 7 Then Exit Sub

Rem CRITICA PARA SABER SE O VALOR DIGITADO FOR MAIOR QUE O SALDO

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
mfl_grid.Col = 1
rs.AddNew "CHAPA", mfl_grid.Text
mfl_grid.Col = 2
rs.Fields("FUNCIONARIO").Value = IIf(IsNull(mfl_grid.Text), " ", Mid$(mfl_grid.Text, 1, 30))
rs.Fields("DT_EVENTO").Value = Format(Now(), "dd/mm/yyyy")
rs.Fields("HORA").Value = Format(Now(), "hh:mm")
rs.Fields("REFERENCIA").Value = "1.00"
mfl_grid.Col = 5
rs.Fields("VALOR").Value = IIf(IsNull(mfl_grid.Text), " ", CDbl(Trim(mfl_grid.Text)))
rs.Fields("TIPO").Value = 1
mfl_grid.Col = 3
rs.Fields("SALDO_ANT").Value = IIf(IsNull(mfl_grid.Text), " ", CDbl(Trim(mfl_grid.Text)))
rs.Update
       
Call CCTempneUniMvFun.MovFuncionario_Cancelar(cMesAno, _
                                             "0", _
                                             cColigada, _
                                             rs)
       
Me.MousePointer = vbDefault

MsgBox "Cancelamento realizado com sucesso!"

Call Limpar_Grid
Call Carregar_Grid
Me.cmd_cancelamento.Enabled = False
Me.txtNome.Text = ""

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmdfechar_Click()
ccodigo_pesquisa = ""
cnome = ""
nValordesc = 0
nValorSaldo = 0

Me.Hide
End Sub

'Private Sub cmdSelecionar_Click()
'
'Me.mfl_grid.Col = 1: ccodigo_pesquisa = Me.mfl_grid.Text
'Me.mfl_grid.Col = 3: nValordesc = Me.mfl_grid.Text
'Me.mfl_grid.Col = 5: nValorSaldo = Me.mfl_grid.Text
'Me.mfl_grid.Col = 2: cnome = Me.mfl_grid.Text
'Me.Hide
'End Sub

Private Sub Form_Activate()
   txtNome.SetFocus
   nTeclou_Enter = 0
End Sub

Private Sub Form_Load()
   Call Carrega_januspesquisa
End Sub
Function Carrega_januspesquisa()

Dim nx As Double
Dim nLinhas As Double
Dim sClass As String

On Error GoTo Erro

Me.mfl_grid.Visible = False
Me.MousePointer = vbHourglass

Call Limpar_Grid
Call Carregar_Grid

Me.mfl_grid.Row = Me.mfl_grid.Rows - 1
If Me.mfl_grid.Rows > 2 Then
   Me.mfl_grid.Col = 2
   If Len(Trim(Me.mfl_grid.Text)) = 0 Then Me.mfl_grid.RemoveItem (Me.mfl_grid.Rows - 1)
End If
Me.mfl_grid.Row = 1

Me.MousePointer = vbDefault
Me.mfl_grid.Visible = True

Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault
   
End Function

Private Sub Limpar_Grid()
Dim nx As Double
Dim nLinhas As Double
Dim nLinhas1 As Double

Me.mfl_grid.Visible = False
mfl_grid.Clear
nLinhas = mfl_grid.Rows

If mfl_grid.Rows > 2 Then
   For nx = mfl_grid.Rows To nLinhas1 - 2 Step -1
       If nx > 2 Then mfl_grid.RemoveItem (nx)
   Next
End If

mfl_grid.Row = 0
Call Ajuste_Tela
Me.txtlidos.Text = 0
Me.mfl_grid.Visible = True

End Sub
Private Sub Grid1_DblClick()

Me.mfl_grid.Col = 1: ccodigo_pesquisa = Me.mfl_grid.Text
Me.mfl_grid.Col = 3: nValordesc = Me.mfl_grid.Text
Me.mfl_grid.Col = 5: nValorSaldo = Me.mfl_grid.Text
Me.mfl_grid.Col = 2: cnome = Me.mfl_grid.Text
Me.Hide
End Sub

Private Sub mfl_grid_Click()
Me.mfl_grid.Col = 2
Me.txtNome.Text = Me.mfl_grid.Text
End Sub

Private Sub Opt_Nome_Click()
Call Carrega_januspesquisa
End Sub

Private Sub Opt_secao_Click()
Call Carrega_januspesquisa
End Sub

Private Sub txtNome_Change()

If Len(Trim(Me.txtNome.Text)) > 0 Then
   Me.cmd_cancelamento.Enabled = True
Else
   Me.cmd_cancelamento.Enabled = False
End If

End Sub

Public Function Carregar_Grid()

Dim x As Variant
Dim nValor As Double
Dim nSaldo As Double
Dim cRec As ADODB.Recordset
Dim nx As Integer

On Error GoTo Erro

Rem ###################################################################
Rem abaixo será realizada a consuta da existencia da movimentacao ref. a
Rem coligada e mes ano de processamento e ter sido digitada
Rem ###################################################################

Set cRec = New ADODB.Recordset
Set cRec = CCTempneUniMvFun.MovFuncionario_ConsMovDigitado(cColigada, cMesAno)

If cRec.RecordCount = 0 Then
   Exit Function
End If

Me.txtlidos.Text = cRec.RecordCount

Me.mfl_grid.Visible = False
mfl_grid.Row = 0
mfl_grid.HighLight = False
Call Ajuste_Tela
mfl_grid.Row = 1

cRec.MoveFirst

While Not cRec.EOF
    
    Set rs = New ADODB.Recordset
    Rem consultar o funcionario na Rm para saber seu salario e calcular seu pagamento
    Set rs = CCTempneUniMvFun.RMFuncionario_Consulta(cColigada, _
                                                     Trim(cRec!MFU_CHAPA))
    
    mfl_grid.Col = 0: mfl_grid.Text = " "
    mfl_grid.Col = 1: mfl_grid.Text = cRec!MFU_CHAPA
    mfl_grid.Col = 7: mfl_grid.Text = " "
    
    If rs.RecordCount = 0 Then
       mfl_grid.Col = 0: mfl_grid.Text = "*"
       mfl_grid.Col = 2: mfl_grid.Text = "Chapa não encontrada"
       mfl_grid.Col = 7: mfl_grid.Text = "Chapa não encontrada"
       GoTo PROXIMO
    Else
       mfl_grid.Col = 2: mfl_grid.Text = rs!NOME
    End If
    
    nSaldo = 0
    If Not IsNull(cRec!SAL_SALDO) Then nSaldo = cRec!SAL_SALDO
    mfl_grid.Col = 3: mfl_grid.Text = Format(nSaldo, "0.00")
    mfl_grid.Col = 4: mfl_grid.Text = "0.00"
    mfl_grid.Col = 5: mfl_grid.Text = Format(cRec!MFU_VALOR, "0.00")
    mfl_grid.Col = 6: mfl_grid.Text = Format(cRec!SAL_SALDO - cRec!MFU_VALOR, "0.00")
    
PROXIMO:
    mfl_grid.Rows = mfl_grid.Rows + 1
    mfl_grid.Row = mfl_grid.Row + 1
    cRec.MoveNext
Wend


Set cRec = Nothing

Me.mfl_grid.Visible = True

Exit Function

Erro:

Me.mfl_grid.Visible = True
Call Limpar_mfl_grid
Call Ajuste_Tela
Set cRec = Nothing

MsgBox "Erro não localizado, anote o numero e chame o responsável. Numero = " & Err.Number
Me.MousePointer = vbDefault

End Function


Private Sub Limpar_mfl_grid()
Dim nx As Double
Dim nLinhas As Double
Dim nLinhas1 As Double

Me.mfl_grid.Visible = False
mfl_grid.Clear
nLinhas = mfl_grid.Rows

If mfl_grid.Rows > 2 Then
   For nx = mfl_grid.Rows To nLinhas1 - 2 Step -1
       If nx > 2 Then mfl_grid.RemoveItem (nx)
   Next
End If

mfl_grid.Row = 0
Call Ajuste_Tela
Me.txtlidos.Text = 0
Me.mfl_grid.Visible = True


End Sub
Private Sub Ajuste_Tela()

mfl_grid.Col = 0: mfl_grid.ColWidth(0) = 400: mfl_grid.Text = "ST"
mfl_grid.Col = 1: mfl_grid.ColWidth(1) = 900: mfl_grid.Text = "CHAPA"
mfl_grid.Col = 2: mfl_grid.ColWidth(2) = 3400: mfl_grid.Text = "FUNCIONARIO"
mfl_grid.Col = 3: mfl_grid.ColWidth(3) = 1400: mfl_grid.Text = "SALDO ANT"
mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 1400: mfl_grid.Text = "VL.UNIMED"
mfl_grid.Col = 5: mfl_grid.ColWidth(5) = 1400: mfl_grid.Text = "DESCONTO"
mfl_grid.Col = 6: mfl_grid.ColWidth(6) = 1400: mfl_grid.Text = "VL.SALDO"
mfl_grid.Col = 7: mfl_grid.ColWidth(7) = 1400: mfl_grid.Text = "SALARIO"
mfl_grid.Col = 8: mfl_grid.ColWidth(8) = 5000: mfl_grid.Text = "OBS"
mfl_grid.Col = 0: mfl_grid.BackColor = &H80FFFF

mfl_grid.Row = 0

mfl_grid.HighLight = False
mfl_grid.ColAlignment(0) = flexAlignCenterCenter
mfl_grid.ColAlignment(1) = flexAlignLeftCenter
mfl_grid.ColAlignment(2) = flexAlignLeftCenter
mfl_grid.ColAlignment(3) = flexAlignRightCenter
mfl_grid.ColAlignment(4) = flexAlignRightCenter
mfl_grid.ColAlignment(5) = flexAlignRightCenter
mfl_grid.ColAlignment(6) = flexAlignRightCenter
mfl_grid.ColAlignment(7) = flexAlignRightCenter
mfl_grid.ColAlignment(8) = flexAlignLeftCenter

End Sub

