VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmGIFIUniCalculoUnimed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculo mensal"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   11310
   Begin VB.Frame frm_arquivo 
      Caption         =   "Selecione a empresa pa o calculo"
      Height          =   1245
      Left            =   30
      TabIndex        =   6
      Top             =   0
      Width           =   11175
      Begin VB.CheckBox chk_saldo_zero 
         Caption         =   "No calculo,zerar os saldos dos demitidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   7410
         TabIndex        =   12
         ToolTipText     =   "Caso esteja marcado esta opção, será repassado o valor de desconto para o proximo mês."
         Top             =   660
         Width           =   2505
      End
      Begin VB.CheckBox chk_repasse 
         Caption         =   "Repassar valor de desconto para o proximo mês."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   7410
         TabIndex        =   11
         ToolTipText     =   "Caso esteja marcado esta opção, será repassado o valor de desconto para o proximo mês."
         Top             =   120
         Width           =   2895
      End
      Begin VB.CommandButton cmd_conf_arquivo 
         BackColor       =   &H0000C000&
         Height          =   555
         Left            =   10500
         MaskColor       =   &H8000000F&
         Picture         =   "frmGIFIUniCalculoUnimed.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Confirmar calculo para atualização da importação"
         Top             =   570
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
         TabIndex        =   7
         Top             =   210
         Width           =   6855
      End
      Begin VB.Label lbl_Msg_Fechamento 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ESTE CALCULO SERÁ REALIZADO PARA O PERIODO DE 01/2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   90
         TabIndex        =   9
         Top             =   810
         Width           =   6915
      End
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   285
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   3
      Top             =   4950
      Width           =   1005
   End
   Begin VB.CommandButton cmd_Atualizar 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Atualizar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   8610
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4950
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   9930
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4950
      Width           =   1275
   End
   Begin VB.CommandButton cmd_imprime_mov 
      BackColor       =   &H00C0C0C0&
      Height          =   330
      Left            =   8040
      Picture         =   "frmGIFIUniCalculoUnimed.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir os Pallet's encontrados"
      Top             =   4950
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSFlexGridLib.MSFlexGrid mfl_grid 
      Height          =   3585
      Left            =   30
      TabIndex        =   4
      Top             =   1290
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   6324
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
   Begin ComctlLib.ProgressBar pbar_processo 
      Height          =   225
      Left            =   2400
      TabIndex        =   10
      Top             =   4980
      Visible         =   0   'False
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label4 
      Caption         =   "Total registros : "
      Height          =   225
      Left            =   60
      TabIndex        =   5
      Top             =   4980
      Width           =   1185
   End
End
Attribute VB_Name = "frmGIFIUniCalculoUnimed"
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
Call Limpar_mfl_grid
Call Confirmar_Dados_coligada
End Sub

Private Sub cbo_coligada_Click()
Call Limpar_mfl_grid
Call Confirmar_Dados_coligada
End Sub

Private Sub cmd_Atualizar_Click()
Dim nx As Integer
Dim nValor As Double
Dim sStatus As String
Dim rs As ADODB.Recordset
Dim RESPOSTA As Integer

On Error GoTo Erro

sStatus = " "

If mfl_grid.Rows > 0 Then
   nx = 1
   While nx < Val(Me.txtlidos.Text) + 1
       mfl_grid.Row = nx
       mfl_grid.Col = 0
       If mfl_grid.Text = "*" Then
          sStatus = "*"
       End If
       nx = nx + 1
   Wend
End If

If sStatus = " " Then
   RESPOSTA = MsgBox("Confirma Atualização dos descontos a serem realizados neste Mês ? ", 20, "Sim/Não?")
Else
   RESPOSTA = MsgBox("Há restrições. Confirma a atualização dos descontos a serem realizados neste Mês ? ", 20, "Sim/Não?")
End If

If RESPOSTA = 7 Then Exit Sub

Me.MousePointer = vbHourglass

Set rs = New ADODB.Recordset

rs.Fields.Append "CHAPA", ADODB.DataTypeEnum.adVarChar, 6
rs.Fields.Append "DESCONTO", ADODB.DataTypeEnum.adDouble
rs.Fields.Append "SALDO", ADODB.DataTypeEnum.adDouble
rs.Fields.Append "SITUACAO", ADODB.DataTypeEnum.adVarChar, 1

rs.Open

mfl_grid.Row = 0

If mfl_grid.Rows > 0 Then
   nx = 1
   While nx < Val(Me.txtlidos.Text) + 1
       mfl_grid.Row = nx
       If Me.chk_repasse.Value = 0 Then
          mfl_grid.Col = 1: rs.AddNew "CHAPA", IIf(Len(Trim(mfl_grid.Text)) = 0, " ", Format(mfl_grid.Text, "000000"))
          mfl_grid.Col = 5: rs.Fields("DESCONTO").Value = IIf(Len(Trim(mfl_grid.Text)) = 0, " ", CDbl(Trim(mfl_grid.Text)))
          mfl_grid.Col = 6: rs.Fields("SALDO").Value = IIf(Len(Trim(mfl_grid.Text)) = 0, " ", CDbl(Trim(mfl_grid.Text)))
          mfl_grid.Col = 0: rs.Fields("SITUACAO").Value = IIf(Len(Trim(mfl_grid.Text)) = 0, " ", Trim(mfl_grid.Text))
          rs.Update
       Else
          mfl_grid.Col = 1: rs.AddNew "CHAPA", IIf(Len(Trim(mfl_grid.Text)) = 0, " ", Mid$(mfl_grid.Text, 1, 6))
          mfl_grid.Col = 5: rs.Fields("DESCONTO").Value = IIf(Len(Trim(mfl_grid.Text)) = 0, " ", CDbl(Trim(mfl_grid.Text)))
          mfl_grid.Col = 6: rs.Fields("SALDO").Value = IIf(Len(Trim(mfl_grid.Text)) = 0, " ", CDbl(Trim(mfl_grid.Text)))
          mfl_grid.Col = 0: rs.Fields("SITUACAO").Value = IIf(Len(Trim(mfl_grid.Text)) = 0, " ", Trim(mfl_grid.Text))
          rs.Update
       End If
       mfl_grid.Col = 1
       nx = nx + 1
   Wend
   
   If rs.RecordCount > 0 Then
      Call CCTempneUniMvFun.MovFuncionario_calcular(sAnoMes, _
                                                    cColigada, _
                                                    rs)
      MsgBox "Calculo Gerado com sucesso!"
      Call Limpar_mfl_grid
      Me.cmd_conf_arquivo.Enabled = False
   Else
      MsgBox "Sem Movimentação, para processar."
   End If
Else
   MsgBox "Sem Movimentação, Retorne."
   Me.MousePointer = vbDefault
   Exit Sub
End If


Me.MousePointer = vbDefault


Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmd_conf_arquivo_Click()

Dim RESPOSTA As Integer
Dim x As Variant
Dim nValor As Double
Dim nSaldo As Double
Dim rs As ADODB.Recordset
Dim nx As Integer

On Error GoTo Erro

'Me.frm_procura_destino.Visible = False

RESPOSTA = MsgBox("'Confirma calculo do arquivo de importação da UNIMED ? ", 20, "Sim/Não?")

cColigada = Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex)

If RESPOSTA = 6 Then
    
    Call Limpar_mfl_grid

    '-------------------------------------> Localizando o Banco de Dados

Rem ###################################################################
Rem abaixo será realizada a consuta da existencia da movimentacao ref. a coligada e mes ano de processamento
Rem ###################################################################
    Set cRec = New ADODB.Recordset
    Set cRec = CCTempneUniMvFun.MovFuncionario_CalculoMensal(Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex), _
                                                             sAnoMes)
   
    
    If cRec.RecordCount = 0 Then
       MsgBox "Movimento não encontrado, procure o responsável."
       Exit Sub
    End If
    
    Me.txtlidos.Text = cRec.RecordCount
    Me.cmd_imprime_mov.Enabled = True
    Me.cmd_Atualizar.Enabled = True

    Me.mfl_grid.Visible = False
    mfl_grid.Row = 0
    mfl_grid.HighLight = False
    Call Ajuste_Tela
    mfl_grid.Row = 1
    Me.pbar_processo.Visible = True
    Me.pbar_processo.Min = 1
    Me.pbar_processo.Max = cRec.RecordCount
    nx = 0
    
    cRec.MoveFirst
    
    While Not cRec.EOF
        
        Set rs = New ADODB.Recordset
        Rem consultar o funcionario na Rm para saber seu salario e calcular seu pagamento
        Set rs = CCTempneUniMvFun.RMFuncionario_Consulta(Str(Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex)), _
                                                         Trim(cRec!MFU_CHAPA))
        
        If rs.RecordCount = 0 Then
           mfl_grid.Col = 0: mfl_grid.Text = "*"
           mfl_grid.Col = 7: mfl_grid.Text = "Chapa não encontrada"
           mfl_grid.Col = 2: mfl_grid.Text = "Chapa não encontrada"
           GoTo PROXIMO
        Else
           mfl_grid.Col = 2: mfl_grid.Text = rs!NOME
        End If
        
        mfl_grid.Col = 0: mfl_grid.Text = " "
        
        Rem modificado aqui para aumntar o tamanho para 7 posicoes em 13/08/2010
        
        mfl_grid.Col = 1: mfl_grid.Text = Format(cRec!MFU_CHAPA, "000000")
        mfl_grid.Col = 7: mfl_grid.Text = " "
        
        nSaldo = 0
        If Not IsNull(cRec!SAL_SALDO) Then nSaldo = cRec!SAL_SALDO
        mfl_grid.Col = 3: mfl_grid.Text = Format(nSaldo, "0.00")
        mfl_grid.Col = 4: mfl_grid.Text = Format(cRec!MFU_VALOR, "0.00")
        
        Rem SALARIO
        mfl_grid.Col = 7: mfl_grid.Text = mfl_grid.Text & Format(rs!SALARIO, "#,##0.00")
        
        Rem ###################################################################
        Rem abaixo sera CALCULADO O VALOR A SER DESCONTADO E A DIFERENCA CASO HAJA
        Rem ###################################################################
        
        nValor = (rs!SALARIO * cRec!MFU_PER_DESC) / 100
        
        If (cRec!MFU_VALOR + nSaldo) <= nValor Then
           mfl_grid.Col = 5: mfl_grid.Text = Format(cRec!MFU_VALOR + nSaldo, "0.00")
           mfl_grid.Col = 6: mfl_grid.Text = "0.00"
        Else
           mfl_grid.Col = 5: mfl_grid.Text = Format(nValor, "0.00")
           mfl_grid.Col = 6: mfl_grid.Text = Format((cRec!MFU_VALOR + nSaldo) - nValor, "0.00")
        End If
        
        
        If Me.chk_repasse.Value = 1 Then
           Rem ###################################################################
           Rem caso o usuario escolha descontar o valor para o proximo mes
           Rem ###################################################################
           mfl_grid.Col = 5: mfl_grid.Text = Format(0, "0.00")
           mfl_grid.Col = 6: mfl_grid.Text = Format((cRec!MFU_VALOR + nSaldo), "0.00")
           mfl_grid.Col = 8: mfl_grid.Text = mfl_grid.Text & "Valor repasado para o proximo mes"
        Else
           Rem ###################################################################
           Rem abaixo sera verificado SITUACAO DO FUNCIONARIO
           Rem ###################################################################
           If rs!CODSITUACAO <> "A" And rs!CODSITUACAO <> "F" Then
              mfl_grid.Col = 0: mfl_grid.Text = rs!CODSITUACAO
              If rs!CODSITUACAO = "D" Then
                 mfl_grid.Col = 8: mfl_grid.Text = mfl_grid.Text & "Demitido, em " & Format(rs!DATADEMISSAO, "DD/MM/YYYY")
                 If Me.chk_saldo_zero.Value = 0 Then
                    mfl_grid.Col = 5: mfl_grid.Text = "0.00"
                    mfl_grid.Col = 6: mfl_grid.Text = Format(cRec!MFU_VALOR + nSaldo, "0.00")
                 Else
                    mfl_grid.Col = 5: mfl_grid.Text = Format(cRec!MFU_VALOR + nSaldo, "0.00")
                    mfl_grid.Col = 6: mfl_grid.Text = "0.00"
                 End If
              ElseIf rs!CODSITUACAO = "I" Then
                 mfl_grid.Col = 8: mfl_grid.Text = mfl_grid.Text & "Aposentado inv., em " & IIf(IsNull(rs!DTAPOSENTADORIA), " ", Format(rs!DTAPOSENTADORIA, "DD/MM/YYYY"))
                 mfl_grid.Col = 5: mfl_grid.Text = "0.00"
                 mfl_grid.Col = 6: mfl_grid.Text = Format(cRec!MFU_VALOR + nSaldo, "0.00")
              ElseIf rs!CODSITUACAO = "O" Then
                 mfl_grid.Col = 8: mfl_grid.Text = mfl_grid.Text & "Doenca ocupacional."
                 mfl_grid.Col = 5: mfl_grid.Text = "0.00"
                 mfl_grid.Col = 6: mfl_grid.Text = Format(cRec!MFU_VALOR + nSaldo, "0.00")
              ElseIf rs!CODSITUACAO = "P" Then
                 mfl_grid.Col = 8: mfl_grid.Text = mfl_grid.Text & "Funcionário em Previdencia."
                 mfl_grid.Col = 5: mfl_grid.Text = "0.00"
                 mfl_grid.Col = 6: mfl_grid.Text = Format(cRec!MFU_VALOR + nSaldo, "0.00")
              ElseIf rs!CODSITUACAO = "T" Then
                 mfl_grid.Col = 8: mfl_grid.Text = mfl_grid.Text & "Afastado em acidente trabalho."
                 mfl_grid.Col = 5: mfl_grid.Text = "0.00"
                 mfl_grid.Col = 6: mfl_grid.Text = Format(cRec!MFU_VALOR + nSaldo, "0.00")
              ElseIf rs!CODSITUACAO = "L" Then
                 mfl_grid.Col = 8: mfl_grid.Text = mfl_grid.Text & "Licenca sem vencimento."
                 mfl_grid.Col = 5: mfl_grid.Text = "0.00"
                 mfl_grid.Col = 6: mfl_grid.Text = Format(cRec!MFU_VALOR + nSaldo, "0.00")
'              ElseIf rs!CODSITUACAO = "F" Then
'                 mfl_grid.Col = 8: mfl_grid.Text = mfl_grid.Text & "Funcionário de ferias"
              Else
                 mfl_grid.Col = 8: mfl_grid.Text = mfl_grid.Text & rs!CODSITUACAO & " - cod. afastamento nao encontrado."
                 mfl_grid.Col = 5: mfl_grid.Text = "0.00"
                 mfl_grid.Col = 6: mfl_grid.Text = Format(cRec!MFU_VALOR + nSaldo, "0.00")
              End If
           End If
           
           Rem aqui marcos 14/02/2014, antes If cRec!MFU_TIPO = 1 Then
           If cRec!MFU_TIPO <> 0 Then
              mfl_grid.Col = 4: mfl_grid.Text = "0.00"
              mfl_grid.Col = 5: mfl_grid.Text = Format(cRec!MFU_VALOR, "0.00")
              mfl_grid.Col = 6: mfl_grid.Text = Format(nSaldo - cRec!MFU_VALOR, "0.00")
              mfl_grid.Col = 8: mfl_grid.Text = mfl_grid.Text & " Valor Digitado."
              Rem ver aqui marcos
           End If
        End If
        

PROXIMO:
        mfl_grid.Rows = mfl_grid.Rows + 1
        mfl_grid.Row = mfl_grid.Row + 1
        nx = nx + 1
        Me.pbar_processo.Value = nx
        cRec.MoveNext
    Wend
  
End If

Me.pbar_processo.Visible = False

Me.mfl_grid.Visible = True

Exit Sub

Erro:

Me.mfl_grid.Visible = True
Call Limpar_mfl_grid
Call Ajuste_Tela
Me.pbar_processo.Visible = False
MsgBox "Erro não localizado, anote o numero e chame o responsável. Numero = " & Err.Number
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
rs.Fields.Append "SALARIO", ADODB.DataTypeEnum.adDouble
rs.Fields.Append "OBS", ADODB.DataTypeEnum.adVarChar, 50

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
       mfl_grid.Col = 4: rs.Fields("SALDO_ANT").Value = IIf(IsNull(mfl_grid.Text), " ", CDbl(Trim(mfl_grid.Text)))
       mfl_grid.Col = 5: rs.Fields("DESCONTO").Value = IIf(IsNull(mfl_grid.Text), " ", CDbl(Trim(mfl_grid.Text)))
       mfl_grid.Col = 6: rs.Fields("SALDO_ATU").Value = IIf(IsNull(mfl_grid.Text), " ", CDbl(Trim(mfl_grid.Text)))
       mfl_grid.Col = 3: rs.Fields("SALARIO").Value = IIf(IsNull(mfl_grid.Text), " ", CDbl(Trim(mfl_grid.Text)))
       mfl_grid.Col = 8: rs.Fields("OBS").Value = IIf(IsNull(mfl_grid.Text), " ", Mid$(mfl_grid.Text, 1, 50))
       rs.Update
   Wend
Else
   MsgBox "Sem Movimentação, Retorne."
   Exit Sub
End If

Set oTela = New frmRelCristalReport

Me.MousePointer = vbHourglass

Set CrystalReport1 = Application.OpenReport(App.Path & "\crptCalculoUnimed.rpt")

CrystalReport1.Database.SetDataSource rs

CrystalReport1.ParameterFields(1).AddCurrentValue "Coligada - " & Me.cbo_coligada.List(Me.cbo_coligada.ListIndex)
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

Private Sub Command1_Click()
Call cmd_Atualizar_Click
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

Call Limpar_mfl_grid
Call carregar_coligada

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
lbl_Msg_Fechamento.Caption = ""

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
Private Sub Ajuste_Tela()

mfl_grid.Col = 0: mfl_grid.ColWidth(0) = 400: mfl_grid.Text = "ST"
mfl_grid.Col = 1: mfl_grid.ColWidth(1) = 900: mfl_grid.Text = "CHAPA"
mfl_grid.Col = 2: mfl_grid.ColWidth(2) = 3400: mfl_grid.Text = "FUNCIONARIO"
mfl_grid.Col = 3: mfl_grid.ColWidth(3) = 1400: mfl_grid.Text = "SALDO ANT"
mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 1400: mfl_grid.Text = "VL.UNIMED"
mfl_grid.Col = 5: mfl_grid.ColWidth(5) = 1400: mfl_grid.Text = "DESCONTO"
mfl_grid.Col = 6: mfl_grid.ColWidth(6) = 1400: mfl_grid.Text = "VL.SALDO"
mfl_grid.Col = 7: mfl_grid.ColWidth(7) = 0: mfl_grid.Text = "SALARIO"
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

Private Sub Confirmar_Dados_coligada()
Dim cRec As ADODB.Recordset

On Error GoTo Erro

Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass
Set cRec = rRec_cliente
Set cRec = CCTempneUniColigada.Coligada_Consultar(sBancoUnimed, Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex))
lbl_Msg_Fechamento.Caption = ""

If cRec.RecordCount > 0 Then
   cRec.MoveFirst
   While Not cRec.EOF
       If cRec!TCO_CODIGO = Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex) Then
          If cRec!TCO_MOV_ABERTO = "1" Then
             lbl_Msg_Fechamento.Caption = "Mov. fechado,precisa abrir para atualizar, Ultimo em " & Mid$(cRec!TCO_ANO_MES_PROC, 5, 2) & "/" & Mid$(cRec!TCO_ANO_MES_PROC, 1, 4) & ", e desconto de " & Format(cRec!TCO_DESCONTO, "0.00") & " %"
             Me.cmd_Atualizar.Visible = False
             Me.cmd_imprime_mov.Visible = False
             Me.cmd_conf_arquivo.Enabled = False
             Me.lbl_Msg_Fechamento.ForeColor = &HFF&
          Else
             lbl_Msg_Fechamento.Caption = "Será calculado dados Para Periodo de " & Mid$(cRec!TCO_ANO_MES_PROC, 5, 2) & "/" & Mid$(cRec!TCO_ANO_MES_PROC, 1, 4) & ", e desconto de " & Format(cRec!TCO_DESCONTO, "0.00") & " %"
             Me.cmd_Atualizar.Visible = True
             Me.cmd_imprime_mov.Visible = True
             Me.cmd_conf_arquivo.Enabled = True
             Me.lbl_Msg_Fechamento.ForeColor = &HC000&
          End If
          sPercentual = cRec!TCO_DESCONTO
          sAnoMes = cRec!TCO_ANO_MES_PROC
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

'''Private Sub Timer1_Timer()
'''If Len(Me.lbl_Msg_Fechamento.Caption) > 0 Then
'''   If Mid$(Me.lbl_Msg_Fechamento.Caption, 1, 6) = "Ultimo" Then
'''      If Me.lbl_Msg_Fechamento.ForeColor = &H80FF& Then
'''         Me.lbl_Msg_Fechamento.ForeColor = &HFF&
'''      Else
'''         Me.lbl_Msg_Fechamento.ForeColor = &H80FF&
'''      End If
'''   Else
'''      If Me.lbl_Msg_Fechamento.ForeColor = &HC000& Then
'''         Me.lbl_Msg_Fechamento.ForeColor = &H80FF80
'''      Else
'''         Me.lbl_Msg_Fechamento.ForeColor = &HC000&
'''      End If
'''   End If
'''Else
'''   Me.lbl_Msg_Fechamento.Visible = False
'''End If
'''
'''End Sub
