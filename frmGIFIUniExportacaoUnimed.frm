VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGIFIUniExportacaoUnimed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Mov. Para a Folha de Pagto"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   10710
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6780
      Top             =   6240
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecione o Mês e Ano para exportar a movimentação"
      Height          =   1515
      Left            =   60
      TabIndex        =   8
      Top             =   90
      Width           =   10545
      Begin VB.CommandButton cmd_conf_arquivo 
         BackColor       =   &H0000C000&
         Height          =   555
         Left            =   9810
         MaskColor       =   &H8000000F&
         Picture         =   "frmGIFIUniExportacaoUnimed.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Verificar existência do movimento fechado para importação"
         Top             =   840
         Width           =   555
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
         TabIndex        =   10
         Top             =   270
         Width           =   9555
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
         TabIndex        =   9
         Text            =   "01/2010"
         Top             =   780
         Width           =   1935
      End
      Begin VB.Label lbl_evento 
         AutoSize        =   -1  'True
         Caption         =   "1"
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
         Left            =   5190
         TabIndex        =   13
         Top             =   900
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código do evento : "
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
         Left            =   2880
         TabIndex        =   12
         Top             =   870
         Width           =   2340
      End
   End
   Begin VB.Frame frm_arquivo 
      Caption         =   "Digite o arquivo para ser exportado"
      Height          =   885
      Left            =   2700
      TabIndex        =   4
      Top             =   4860
      Width           =   7845
      Begin VB.TextBox txt_Arq_Exportacao 
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
         Height          =   405
         Left            =   90
         MaxLength       =   50
         TabIndex        =   5
         Text            =   "C:\ARQ_UNIMED\EXPORTACAO\MovExportacao.TXT"
         ToolTipText     =   "Nome do ae-mail responsável pelo envio da mensagem"
         Top             =   270
         Width           =   7605
      End
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   1230
      MaxLength       =   6
      TabIndex        =   3
      Top             =   4860
      Width           =   1005
   End
   Begin VB.CommandButton cmd_confirmar 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Confirmar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   8040
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5970
      Width           =   1275
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   9390
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5970
      Width           =   1275
   End
   Begin VB.CommandButton cmd_imprime_mov 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   330
      Left            =   7500
      Picture         =   "frmGIFIUniExportacaoUnimed.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir o Movimento de importação para a folha."
      Top             =   5970
      Width           =   465
   End
   Begin MSFlexGridLib.MSFlexGrid mfl_grid 
      Height          =   3105
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   5477
      _Version        =   393216
      Cols            =   8
      AllowBigSelection=   0   'False
      TextStyle       =   3
      TextStyleFixed  =   2
      HighLight       =   2
      ScrollBars      =   2
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
   Begin VB.Label lbl_calculado 
      Caption         =   "Movimento nao calculado ou aberto. verifique!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5940
      Width           =   5625
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10710
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label4 
      Caption         =   "Total registros : "
      Height          =   285
      Left            =   30
      TabIndex        =   7
      Top             =   4890
      Width           =   1185
   End
End
Attribute VB_Name = "frmGIFIUniExportacaoUnimed"
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
Private bCalculado As Boolean ' contera a critica do movimento caso ainda nao esteja calclado
Private Sub cbo_coligada_Change()
Call Limpar_mfl_grid
Call Confirmar_Dados_coligada
Call Confirmar_Meses_Fechados
End Sub

Private Sub cbo_coligada_Click()
Call Limpar_mfl_grid
Call Confirmar_Dados_coligada
End Sub

Private Sub cbo_mes_ano_Change()
sAnoMes = Me.cbo_mes_ano.ItemData(Me.cbo_mes_ano.ListIndex)
Call Confirmar_Dados_coligada
End Sub

Private Sub cbo_mes_ano_Click()
sAnoMes = Me.cbo_mes_ano.ItemData(Me.cbo_mes_ano.ListIndex)
Call Confirmar_Dados_coligada
End Sub

Private Sub cmd_conf_arquivo_Click()
Dim RESPOSTA As Integer
Dim x As Variant
Dim nValor As Double
Dim nSaldo As Double
Dim rs As ADODB.Recordset

On Error GoTo Erro

'Me.frm_procura_destino.Visible = False

RESPOSTA = MsgBox("'Confirma Leitura do Movimento calculado para importação da UNIMED ? ", 20, "Sim/Não?")

cColigada = Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex)

If RESPOSTA = 6 Then

    Me.MousePointer = vbHourglass
    
    Call Limpar_mfl_grid

    '-------------------------------------> Localizando o Banco de Dados

Rem ###################################################################
Rem abaixo será realizada a consuta da existencia da movimentacao ref. a coligada e mes ano de processamento
Rem ###################################################################
    Set cRec = New ADODB.Recordset
    Set cRec = CCTempneUniMvFun.MovFuncionario_ConsMovExpMensal(Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex), _
                                                                sAnoMes)
   
    
    If cRec.RecordCount = 0 Then
       MsgBox "Movimento não encontrado, procure o responsável."
       Me.MousePointer = vbDefault
       Exit Sub
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
        Rem consultar o funcionario na Rm para saber seu nome e se não é demitido
        Set rs = CCTempneUniMvFun.RMFuncionario_Consulta(Str(Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex)), Trim(cRec!MFU_CHAPA))
        
        If rs.RecordCount = 0 Then
           mfl_grid.Col = 0: mfl_grid.Text = "*"
           mfl_grid.Col = 7: mfl_grid.Text = "Chapa não encontrada"
           mfl_grid.Col = 2: mfl_grid.Text = "Chapa não encontrada"
           GoTo PROXIMO
        Else
           If rs!CODSITUACAO = "D" Or cRec!MFU_TIPO = 2 Then GoTo LEITURA 'caso seja demitido ou tipo = 2, não é para descontar na folha 14/02/2014
           mfl_grid.Col = 2: mfl_grid.Text = rs!NOME
        End If
        
        mfl_grid.Col = 0: mfl_grid.Text = " "
        mfl_grid.Col = 1: mfl_grid.Text = cRec!MFU_CHAPA
        mfl_grid.Col = 7: mfl_grid.Text = " "
        
        nSaldo = 0
        If Not IsNull(cRec!SAL_SALDO) Then nSaldo = cRec!SAL_SALDO
        mfl_grid.Col = 3: mfl_grid.Text = "0.00"
        mfl_grid.Col = 4: mfl_grid.Text = Format(cRec!MFU_VALOR, "0.00")
        
        mfl_grid.Col = 5: mfl_grid.Text = Format(cRec!MFU_VALOR_DESC, "0.00")
        mfl_grid.Col = 6: mfl_grid.Text = Format(nSaldo, "0.00")
        

PROXIMO:
        mfl_grid.Rows = mfl_grid.Rows + 1
        mfl_grid.Row = mfl_grid.Row + 1
LEITURA:
        cRec.MoveNext
    Wend
  
End If

Me.txtlidos.Text = mfl_grid.Rows - 1
mfl_grid.Col = 1

If Len(Trim(mfl_grid.Text)) = 0 Then
   mfl_grid.Rows = mfl_grid.Rows - 1
   mfl_grid.Row = mfl_grid.Row - 1
   Me.txtlidos.Text = mfl_grid.Rows - 2
End If

Me.cmd_imprime_mov.Enabled = True
Me.cmd_confirmar.Enabled = True

Me.mfl_grid.Visible = True
Me.MousePointer = vbDefault

Exit Sub

Erro:

MsgBox "Erro não localizado, anote o numero e chame o responsável. Numero = " & Err.Number

Me.MousePointer = vbDefault

End Sub

Private Sub cmd_Confirmar_Click()
Dim sNome_Arq As String
Dim RESPOSTA As Integer
Dim X1 As Double
Dim X2 As Double
Dim Y As Double
Dim Nada As String
Dim x As Variant
Dim nNumero As Double
Dim rs As ADODB.Recordset

On Error GoTo Erro

sNome_Arq = Me.txt_Arq_Exportacao.Text

RESPOSTA = MsgBox("'Confirma Gerar arquivo de Exportação para a Folha de Pagto. ? " & sNome_Arq, 20, "Sim/Não?")

If RESPOSTA = 6 Then
    Me.MousePointer = vbHourglass
    Open sNome_Arq For Random Access Read Write As #11 Len = Len(ArqMovUnimed)
    Close 11
    Kill sNome_Arq
    Open sNome_Arq For Random Access Read Write As #11 Len = Len(ArqMovUnimed)
   
    
    Me.mfl_grid.Visible = False
    mfl_grid.Row = 1
    
    For Y = 1 To Val(Me.txtlidos.Text)
        
        mfl_grid.Col = 1
        If Len(Trim(mfl_grid.Text)) > 0 Then
           ArqMovUnimed.Fchapa = Mid$(Trim(mfl_grid.Text), 2, Len(Trim(mfl_grid.Text)))
           ArqMovUnimed.Fdtpagto = Format(Now(), "ddmmyyyy")
           ArqMovUnimed.Fcodevento = Me.lbl_evento.Caption
           ArqMovUnimed.Fhora = "0" & Format(Now(), "hh:mm")
           ArqMovUnimed.Frefer = "000000000000.00"
           mfl_grid.Col = 5: ArqMovUnimed.Fvalor = Replace(Replace(Format(mfl_grid.Text, "000000000000.00"), ".", ""), ",", ".")
           mfl_grid.Col = 5: ArqMovUnimed.Fvaloreal = Replace(Replace(Format(mfl_grid.Text, "000000000000.00"), ".", ""), ",", ".")
           ArqMovUnimed.Falterado = "N"
           ArqMovUnimed.FFinal = Chr$(13) + Chr$(10)
           X2 = X2 + 1
           Put 11, X2, ArqMovUnimed
        End If
        mfl_grid.Rows = mfl_grid.Rows + 1
        mfl_grid.Row = mfl_grid.Row + 1
        
    Next
    Me.MousePointer = vbDefault
    Close #11
   
End If

Me.mfl_grid.Visible = True

MsgBox "Arquivo gerado com sucesso"

Exit Sub

Erro:

If Err.Number = 52 Then
   MsgBox "Arquivo não encontrado em " + sNome_Arq + "Movimento NÃO pode ser realizado"
Else
   MsgBox "Erro não localizado, anote o numero e chame o responsável. Numero = " & Err.Number
   Close #11
End If

Me.MousePointer = vbDefault

End Sub

Private Sub cmd_imprime_mov_Click()
Dim oTela As frmRelCristalReport
Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim rs As ADODB.Recordset
Dim nx As Integer
Dim nValor As Double
Dim sStatus As String

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set rs = New ADODB.Recordset

rs.Fields.Append "CHAPA", ADODB.DataTypeEnum.adVarChar, 7
rs.Fields.Append "FUNCIONARIO", ADODB.DataTypeEnum.adVarChar, 30
rs.Fields.Append "SALDO_ANT", ADODB.DataTypeEnum.adDouble
rs.Fields.Append "DESCONTO", ADODB.DataTypeEnum.adDouble
rs.Fields.Append "SALDO_ATU", ADODB.DataTypeEnum.adDouble

rs.Open

mfl_grid.Row = 0

'mfl_grid.Col = 1: mfl_grid.ColWidth(1) = 900: mfl_grid.Text = "CHAPA"
'mfl_grid.Col = 2: mfl_grid.ColWidth(2) = 3400: mfl_grid.Text = "FUNCIONARIO"
'mfl_grid.Col = 3: mfl_grid.ColWidth(3) = 1400: mfl_grid.Text = "SALDO ANT"
'mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 0: mfl_grid.Text = "VL.UNIMED"
'mfl_grid.Col = 5: mfl_grid.ColWidth(5) = 1400: mfl_grid.Text = "DESCONTO"
'mfl_grid.Col = 6: mfl_grid.ColWidth(6) = 1400: mfl_grid.Text = "VL.SALDO"



If mfl_grid.Rows > 0 Then
   nx = 0
   While nx < Val(Me.txtlidos.Text)
       nx = nx + 1
       mfl_grid.Row = nx
       mfl_grid.Col = 1: rs.AddNew "CHAPA", IIf(IsNull(mfl_grid.Text), " ", Mid$(mfl_grid.Text, 1, 7))
       mfl_grid.Col = 2: rs.Fields("FUNCIONARIO").Value = IIf(IsNull(mfl_grid.Text), " ", Mid$(mfl_grid.Text, 1, 30))
       mfl_grid.Col = 3: rs.Fields("SALDO_ANT").Value = IIf(IsNull(mfl_grid.Text), " ", CDbl(Trim(mfl_grid.Text)))
       mfl_grid.Col = 5: rs.Fields("DESCONTO").Value = IIf(IsNull(mfl_grid.Text), " ", CDbl(Trim(mfl_grid.Text)))
       mfl_grid.Col = 6: rs.Fields("SALDO_ATU").Value = IIf(IsNull(mfl_grid.Text), " ", CDbl(Trim(mfl_grid.Text)))
       rs.Update
   Wend
Else
   MsgBox "Sem Movimentação, Retorne."
   Exit Sub
End If

Set oTela = New frmRelCristalReport

Me.MousePointer = vbHourglass

Set CrystalReport1 = Application.OpenReport(App.Path & "\crptExportacaoUnimed.rpt")

CrystalReport1.Database.SetDataSource rs

CrystalReport1.ParameterFields(1).AddCurrentValue "Coligada - " & Me.cbo_coligada.List(Me.cbo_coligada.ListIndex) & " / Periodo - " & Me.cbo_mes_ano.List(Me.cbo_mes_ano.ListIndex)
CrystalReport1.ParameterFields(1).DiscreteOrRangeKind = crDiscreteValue

CrystalReport1.ParameterFields(2).AddCurrentValue "teste2"
CrystalReport1.ParameterFields(2).DiscreteOrRangeKind = crDiscreteValue


oTela.CRV_RELATORIO.ReportSource = CrystalReport1
oTela.CRV_RELATORIO.ViewReport

rs.Clone

Me.MousePointer = vbDefault

oTela.Show 0

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
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
'bCalculado = False
End Sub

Private Sub Form_Load()
Dim nx As Integer

Me.Top = 0
Me.Left = 0

Call Limpar_mfl_grid
Call carregar_coligada
Call carregar_Meses_Fechados
Me.txt_Arq_Exportacao.Text = "C:\ARQ_UNIMED\EXPORTACAO\Mov" & Format(Now(), "MMYYYY") & "Exp.TXT"
End Sub
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
Me.cmd_imprime_mov.Enabled = False


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
   MsgBox "Não existem Fechamento das empresas Coligadas, procure o responsável."
End If

Me.MousePointer = vbDefault

Set cRec = Nothing
Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub
Private Sub Ajuste_Tela()

mfl_grid.Col = 0: mfl_grid.ColWidth(0) = 400: mfl_grid.Text = "ST"
mfl_grid.Col = 1: mfl_grid.ColWidth(1) = 900: mfl_grid.Text = "CHAPA"
mfl_grid.Col = 2: mfl_grid.ColWidth(2) = 3400: mfl_grid.Text = "FUNCIONARIO"
mfl_grid.Col = 3: mfl_grid.ColWidth(3) = 0: mfl_grid.Text = "SALDO ANT"
mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 0: mfl_grid.Text = "VL.UNIMED"
mfl_grid.Col = 5: mfl_grid.ColWidth(5) = 1400: mfl_grid.Text = "DESCONTO"
mfl_grid.Col = 6: mfl_grid.ColWidth(6) = 1400: mfl_grid.Text = "VL.SALDO"
mfl_grid.Col = 7: mfl_grid.ColWidth(7) = 0: mfl_grid.Text = "OBS"
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
mfl_grid.ColAlignment(7) = flexAlignLeftCenter

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
'          If cRec!TCO_MOV_ABERTO = "1" Then
''             lbl_Msg_Fechamento.Caption = "Mov. fechado,precisa abrir para atualizar, Ultimo em " & Mid$(cRec!TCO_ANO_MES_PROC, 5, 2) & "/" & Mid$(cRec!TCO_ANO_MES_PROC, 1, 4) & ", e desconto de " & Format(cRec!TCO_DESCONTO, "0.00") & " %"
''             Me.cmd_Atualizar.Visible = False
''             Me.cmd_imprime_mov.Visible = False
''             Me.cmd_conf_arquivo.Enabled = False
''             Me.lbl_Msg_Fechamento.ForeColor = &HFF&
'          Else
''             lbl_Msg_Fechamento.Caption = "Será calculado dados Para Periodo de " & Mid$(cRec!TCO_ANO_MES_PROC, 5, 2) & "/" & Mid$(cRec!TCO_ANO_MES_PROC, 1, 4) & ", e desconto de " & Format(cRec!TCO_DESCONTO, "0.00") & " %"
''             Me.cmd_Atualizar.Visible = True
''             Me.cmd_imprime_mov.Visible = True
''             Me.cmd_conf_arquivo.Enabled = True
''             Me.lbl_Msg_Fechamento.ForeColor = &HC000&
'          End If
          sPercentual = cRec!TCO_DESCONTO
          Me.lbl_evento.Caption = cRec!TCO_VERBA
Rem VERIFICAR SE O MOVIMENTO ESTA FECHADO
          If cRec!TCO_MOV_ABERTO <> "1" Then
             bCalculado = False
          Else
             bCalculado = True
          End If
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



Private Sub Confirmar_Meses_Fechados()
Dim nx As Integer
Dim cRec As ADODB.Recordset

On Error GoTo Erro

Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass
Set cRec = rRec_cliente
Set cRec = CCTempneUniMvFun.MovFuncionario_ConsMesFechado(cColigada)

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

Private Sub Timer1_Timer()
If bCalculado = False Then
    Me.cmd_imprime_mov.Visible = False
    Me.cmd_confirmar.Visible = False
    If Me.lbl_calculado.Visible = True Then
       Me.lbl_calculado.Visible = False
    Else
       Me.lbl_calculado.Visible = True
    End If
Else
    Me.lbl_calculado.Visible = False
    Me.cmd_imprime_mov.Visible = True
    Me.cmd_confirmar.Visible = True
End If
End Sub
