VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGIFIUniConsultarSaldo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Saldos"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10815
   Begin VB.CommandButton cmd_restaurar 
      BackColor       =   &H0000C000&
      Caption         =   "&Restaurar"
      Height          =   330
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "este botão restaurará o movimento que foi salvo antes do ultimo fechamento"
      Top             =   5460
      Width           =   2235
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   1290
      MaxLength       =   6
      TabIndex        =   8
      Top             =   5460
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Height          =   4395
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   10605
      Begin MSFlexGridLib.MSFlexGrid mfl_grid 
         Height          =   4125
         Left            =   90
         TabIndex        =   6
         ToolTipText     =   "Clique duas vezes no funcionário para ver o seu extrato"
         Top             =   180
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7276
         _Version        =   393216
         Cols            =   4
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
         TabIndex        =   7
         Top             =   1410
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecione a empresa/Pesquisar Mov digitado"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   10575
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
         TabIndex        =   4
         Top             =   270
         Width           =   6765
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
         Left            =   210
         TabIndex        =   3
         Text            =   "01/2010"
         Top             =   930
         Width           =   1935
      End
      Begin VB.CommandButton cmd_confirmar_Cons 
         BackColor       =   &H00FF8080&
         Caption         =   "Pesquisar"
         Height          =   465
         Left            =   8790
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "pesquisar movimentos digitados no Periodo/Coligada"
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5460
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Total registros : "
      Height          =   225
      Left            =   90
      TabIndex        =   9
      Top             =   5520
      Width           =   1185
   End
End
Attribute VB_Name = "frmGIFIUniConsultarSaldo"
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
Private rs As ADODB.Recordset
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


Private Sub cmd_confirmar_Cons_Click()

Dim nx As Double
Dim nLinhas As Double


On Error GoTo Erro

Me.mfl_grid.Visible = False
Me.MousePointer = vbHourglass
Set rs = New ADODB.Recordset

Rem VERIFICAR OS FUNCIONARIOS QUE EXISTEM SALDO NO CADASTRO

Set rs = CCTempneUniMvFun.MovFuncionario_Consulta_Saldo(Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex), "")

If rs.RecordCount > 0 Then
   Call Limpar_Grid
   Call Carregar_Grid
End If

Me.MousePointer = vbDefault
Me.mfl_grid.Visible = True

Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault
   
End Sub

Private Sub Limpar_Grid()
Dim nx As Double
Dim nLinhas As Double
Dim nLinhas1 As Double

mfl_grid.Clear
nLinhas = mfl_grid.Rows

If mfl_grid.Rows > 2 Then
   For nx = mfl_grid.Rows To nLinhas1 - 2 Step -1
       If nx > 2 Then mfl_grid.RemoveItem (nx)
   Next
End If

mfl_grid.Row = 0
mfl_grid.Col = 0: mfl_grid.ColWidth(0) = 900:  mfl_grid.Text = "CHAPA"
mfl_grid.Col = 1:  mfl_grid.ColWidth(1) = 4500: mfl_grid.Text = "NOME"
mfl_grid.Col = 2: mfl_grid.ColWidth(2) = 3500: mfl_grid.Text = "SEÇÃO"
mfl_grid.Col = 3: mfl_grid.ColWidth(3) = 1200: mfl_grid.Text = "SALDO"
mfl_grid.Col = 2: mfl_grid.BackColor = &H80FFFF

mfl_grid.Row = 0

mfl_grid.HighLight = False


End Sub
Public Function Carregar_Grid()

Dim nx As Double
Dim nLinhas As String
Dim sClass As String

Me.mfl_grid.Visible = False
Me.MousePointer = vbHourglass

mfl_grid.Row = 1
rs.MoveFirst

For nx = 1 To rs.RecordCount
    Rem ACHAR O FUNCIONARIO NO CADASTRO DE FUNCIONARIOS NA RM
    If rs!SAL_SALDO > 0 Then
       Set cRec = New ADODB.Recordset
       Set cRec = CCTempneTabRegPagto.TabRegPagto_Consultar_Funcionario(sBancoRM, _
                                                                        Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex), _
                                                                        Format(rs!SAL_CHAPA, "00000"))
'       If cRec!CODSITUACAO <> "A" And cRec!CODSITUACAO <> "F" Then
          nLinhas = Format(cRec.Fields("CHAPA"), "000000")
          mfl_grid.Col = 0: mfl_grid.Text = nLinhas
          mfl_grid.Col = 1: mfl_grid.Text = cRec.Fields("nome")
          mfl_grid.Col = 2: mfl_grid.Text = cRec!CODSITUACAO & " - " & cRec.Fields("descricao")
          mfl_grid.Col = 3: mfl_grid.Text = Format(rs.Fields("SAL_SALDO"), "#,##0.00")
          rs.MoveNext
          If Not rs.EOF Then
             mfl_grid.Rows = mfl_grid.Rows + 1
             mfl_grid.Row = mfl_grid.Row + 1
          End If
'       Else
'          rs.MoveNext
'       End If
    Else
       rs.MoveNext
    End If
Next

mfl_grid.Col = 1
mfl_grid.Sort = flexSortStringAscending

If mfl_grid.Rows > 2 Then
   mfl_grid.Rows = mfl_grid.Rows
   If Len(Trim(mfl_grid.Text)) = 0 Then mfl_grid.RemoveItem (mfl_grid.Row)
   mfl_grid.Row = 1
   If Len(Trim(mfl_grid.Text)) = 0 Then mfl_grid.RemoveItem (mfl_grid.Row)
   
End If

Me.mfl_grid.Visible = True
Me.txtlidos.Text = mfl_grid.Rows - 1
Me.MousePointer = vbDefault
Set cRec = Nothing

Exit Function

Error:

Set cRec = Nothing
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


Private Sub cmd_restaurar_Click()
Dim RESPOSTA As Integer

RESPOSTA = MsgBox("Restaurar cópia do Mov. antes do ultimo fechamento?", 20, "Sim/Não?")

On Error Resume Next

If RESPOSTA = 6 Then
   Rem SERÁ REALIZADA UMA COPIA DOS DADOS EM UMA TABELA AUXILIAR.
   Call CCTempneUniMvFun.MovFuncionario_Restaura_Copia
   MsgBox "Copia realizada com sucesso!"
End If

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
Call Limpar_Grid

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

'Function Pesquisar_Saldo()
'On Error GoTo Erro
'
'Me.MousePointer = vbHourglass
'Set cRec = New ADODB.Recordset
'
'Set cRec = CCTempneUniMvFun.MovFuncionario_Consulta_Saldo(Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex), _
'                                                           Me.txtCODFUNC.Text)
'
'If cRec.RecordCount > 0 Then
'   Me.TXT_SALDO_ANT.Text = Format(cRec!SAL_SALDO, "0.00")
'End If
'
'Me.MousePointer = vbDefault
'
'Exit Function
'
'Erro:
'MsgBox Err.Description
'Me.MousePointer = vbDefault
'
'End Function




Private Sub mfl_grid_DblClick()

Dim oTela As frmGIFIUniPesquisarextrato

Set oTela = New frmGIFIUniPesquisarextrato
oTela.cColigada = Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex)
Me.mfl_grid.Col = 0
oTela.ccodigo_pesquisa = Me.mfl_grid.Text
Me.mfl_grid.Col = 1
oTela.Label2 = Me.mfl_grid.Text


oTela.Show 1
Set oTela = Nothing

End Sub
