VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmGIFIUniImportacaoUnimed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Movimentação"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   10710
   Begin VB.Frame frm_procura_destino 
      Height          =   4065
      Left            =   150
      TabIndex        =   4
      Top             =   1290
      Visible         =   0   'False
      Width           =   4965
      Begin VB.DriveListBox drv_destino 
         Height          =   315
         Left            =   60
         TabIndex        =   7
         Top             =   120
         Width           =   4785
      End
      Begin VB.DirListBox dir_destino 
         Height          =   2340
         Left            =   60
         TabIndex        =   6
         Top             =   510
         Width           =   4785
      End
      Begin VB.FileListBox flie_destino 
         Height          =   1065
         Left            =   60
         Pattern         =   "*.txt"
         TabIndex        =   5
         Top             =   2910
         Width           =   4785
      End
   End
   Begin ComctlLib.ProgressBar pbar_processo 
      Height          =   225
      Left            =   2310
      TabIndex        =   16
      Top             =   5460
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmd_imprime_mov 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   345
      Left            =   7380
      Picture         =   "frmGIFIUniImportacaoUnimed.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Imprimir os Pallet's encontrados"
      Top             =   5400
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   1275
   End
   Begin VB.CommandButton cmd_atualizar 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Importar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   8010
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtlidos 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   9
      Text            =   "0"
      Top             =   5430
      Width           =   1005
   End
   Begin VB.Frame frm_arquivo 
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   $"frmGIFIUniImportacaoUnimed.frx":0532
      Top             =   30
      Width           =   10485
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
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   300
         Width           =   9555
      End
      Begin VB.CommandButton cmd_conf_arquivo 
         BackColor       =   &H0000C000&
         Height          =   555
         Left            =   9810
         MaskColor       =   &H8000000F&
         Picture         =   "frmGIFIUniImportacaoUnimed.frx":060D
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Confirmar Arquivo para importação"
         Top             =   1140
         Width           =   555
      End
      Begin VB.CommandButton cmd_achar_arquivo 
         Height          =   345
         Left            =   9090
         Picture         =   "frmGIFIUniImportacaoUnimed.frx":0917
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Procurar arquivo para concistência dos dados."
         Top             =   840
         Width           =   525
      End
      Begin VB.TextBox txt_Arq_Importacao 
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
         MaxLength       =   100
         TabIndex        =   3
         Text            =   "c:\"
         ToolTipText     =   "Arquivo que será analisado para importação dos dados"
         Top             =   810
         Width           =   9555
      End
      Begin VB.Label lbl_Msg_Fechamento 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   150
         TabIndex        =   15
         Top             =   1320
         Width           =   9555
      End
   End
   Begin MSFlexGridLib.MSFlexGrid mfl_grid 
      Height          =   3375
      Left            =   60
      TabIndex        =   8
      Top             =   1950
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   5953
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
   Begin VB.Label Label4 
      Caption         =   "Total registros : "
      Height          =   225
      Left            =   30
      TabIndex        =   12
      Top             =   5460
      Width           =   1185
   End
End
Attribute VB_Name = "frmGIFIUniImportacaoUnimed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Private sAnoMes As String 'contem o ano mes referente ao arquivo da coligada
Private sPercentual As String 'contem o percentual referente ao arquivo da coligada
Private sMovSituacao As String 'cotem a situacao do arquivo da coligada 0=aberto, 1=fechado

Private Sub cbo_coligada_Change()
Call Limpar_mfl_grid
Call Confirmar_Dados_coligada
End Sub

Private Sub cbo_coligada_Click()
Call Limpar_mfl_grid
Call Confirmar_Dados_coligada
End Sub

Private Sub cmd_achar_arquivo_Click()
On Error GoTo Erro

If Me.frm_procura_destino.Visible = False Then
   Me.frm_procura_destino.Visible = True
   Me.drv_destino = "C:\"
   Me.dir_destino = "C:\ARQ_UNIMED\IMPORTACAO"
Else
   Me.frm_procura_destino.Visible = False
End If
Exit Sub

Erro:

If Err.Number = 76 Then
   MsgBox "Diretorio C:\ARQ_UNIMED\IMPORTACAO, não existe, Crie o novo diretório"
   Me.dir_destino = "C:\"
End If
End Sub

Private Sub cmd_Atualizar_Click()

Dim nx As Integer
Dim nValor As Double
Dim sStatus As String
Dim rs As ADODB.Recordset
Dim RESPOSTA As Integer

On Error GoTo Erro

Rem verificando se há restricoes

mfl_grid.Row = 0

If Val(Me.txtlidos.Text) = 0 Then
   MsgBox "Não há resgistros para ser processado."
   Exit Sub
End If

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
   RESPOSTA = MsgBox("Confirma carga do arquivo de importação da UNIMED ? ", 20, "Sim/Não?")
Else
   RESPOSTA = MsgBox("Há restrições. Caso confirme a importação será realizada apenas os válidos. Confirma ? ", 20, "Sim/Não?")
End If

If RESPOSTA = 7 Then Exit Sub

Me.MousePointer = vbHourglass

Rem SERÁ REALIZADA UMA COPIA DOS DADOS EM UMA TABELA AUXILIAR.

Call CCTempneUniMvFun.MovFuncionario_Copia_fechamento

Set rs = New ADODB.Recordset

rs.Fields.Append "CHAPA", ADODB.DataTypeEnum.adVarChar, 7
rs.Fields.Append "FUNCIONARIO", ADODB.DataTypeEnum.adVarChar, 30
rs.Fields.Append "DT_EVENTO", ADODB.DataTypeEnum.adVarChar, 10
rs.Fields.Append "HORA", ADODB.DataTypeEnum.adVarChar, 5
rs.Fields.Append "REFERENCIA", ADODB.DataTypeEnum.adDouble
rs.Fields.Append "VALOR", ADODB.DataTypeEnum.adDouble
rs.Fields.Append "TIPO", ADODB.DataTypeEnum.adInteger

rs.Open

mfl_grid.Row = 0

If mfl_grid.Rows > 0 Then
   nx = 1
   While nx < Val(Me.txtlidos.Text) + 1
       mfl_grid.Row = nx
       mfl_grid.Col = 0
       If mfl_grid.Text <> "*" Then
          mfl_grid.Col = 1: rs.AddNew "CHAPA", IIf(IsNull(mfl_grid.Text), " ", Mid$(mfl_grid.Text, 1, 7))
          mfl_grid.Col = 2: rs.Fields("FUNCIONARIO").Value = IIf(IsNull(mfl_grid.Text), " ", Mid$(mfl_grid.Text, 1, 30))
          mfl_grid.Col = 3: rs.Fields("DT_EVENTO").Value = IIf(IsNull(mfl_grid.Text), " ", Mid$(mfl_grid.Text, 1, 10))
          mfl_grid.Col = 4: rs.Fields("HORA").Value = "00:00" 'IIf(IsNull(mfl_grid.Text), " ", Mid$(mfl_grid.Text, 2, 5))
          mfl_grid.Col = 5: rs.Fields("REFERENCIA").Value = IIf(IsNull(mfl_grid.Text), " ", CDbl(Trim(mfl_grid.Text)))
          mfl_grid.Col = 6: rs.Fields("VALOR").Value = IIf(IsNull(mfl_grid.Text), " ", CDbl(Trim(mfl_grid.Text)))
          rs.Fields("TIPO").Value = 0
          rs.Update
       End If
       nx = nx + 1
   Wend
   
   If rs.RecordCount > 0 Then
      Call CCTempneUniMvFun.MovFuncionario_Incluir(sAnoMes, _
                                                   sPercentual, _
                                                   Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex), _
                                                   rs)
      MsgBox "Movimentação atualizada com Sucesso."
      Me.cmd_atualizar.Visible = False
   Else
      MsgBox "Sem Movimentação, para processar."
   End If
Else
   MsgBox "Sem Movimentação, Retorne."
   Exit Sub
End If

Me.MousePointer = vbDefault


Exit Sub

Erro:

MsgBox "chapa- " & rs!CHAPA & " Func - " & rs!FUNCIONARIO & " Dt- " & rs!DT_EVENTO & " hora - " & rs!HORA & " Ref- " & rs!REFERENCIA & "Vl - " & rs!VALOR

MsgBox Err.Description, , Me.Caption

Me.MousePointer = vbDefault

End Sub

Private Sub cmd_conf_arquivo_Click()

Dim sNome_Arq As String
Dim RESPOSTA As Integer
Dim X1 As Double
Dim X2 As Double
Dim Y As Double
Dim Nada As String
Dim x As Variant
Dim nNumero As Double
Dim rs As ADODB.Recordset
Dim schapa As String
Dim ddata_Compara As Date


On Error GoTo Erro

Me.frm_procura_destino.Visible = False

If Me.cbo_coligada.ListIndex = -1 Then
   MsgBox "Selecione a empresa coligada"
   Me.cbo_coligada.SetFocus
   Exit Sub
End If

sNome_Arq = Me.txt_Arq_Importacao.Text

RESPOSTA = MsgBox("'Confirma carga do arquivo de importação da UNIMED ? " & sNome_Arq, 20, "Sim/Não?")

If RESPOSTA = 6 Then
    
    Call Limpar_mfl_grid
    Me.Refresh
    
    '-------------------------------------> Localizando o Banco de Dados

    Nada = Dir$(sNome_Arq)
    
    If Nada = "" Or Nada = ".rnd" Then
       MsgBox "Arquivo não encontrado em " + sNome_Arq + ". Carga NÃO pode ser feita"
       Exit Sub
    End If

    Open sNome_Arq For Random Access Read Write As #11 Len = Len(ArqMovUnimed)
   
    X2 = 1
    Get 11, X2, ArqMovUnimed
    While Format(Val(Trim(ArqMovUnimed.Fchapa)), "000000") > 0
          X2 = X2 + 1
          If X2 = 330 Then
             X2 = X2
          End If
          Get 11, X2, ArqMovUnimed
    Wend
       
'       Close 11
'       Exit Sub
'    End If
    
    X1 = X2 - 1
    X2 = 0
    
    If X1 > 0 Then
       Me.txtlidos.Text = X1
       Me.cmd_imprime_mov.Visible = True
       Me.cmd_imprime_mov.Enabled = True
       Me.cmd_atualizar.Visible = True
       Me.cmd_atualizar.Enabled = True
    Else
       Me.cmd_atualizar.Visible = False
       Exit Sub
    End If
    
    Me.mfl_grid.Visible = False
    mfl_grid.Row = 0
    mfl_grid.HighLight = False
    Call Ajuste_Tela
    mfl_grid.Row = 1
    Me.SetFocus
    Me.pbar_processo.Min = 0
    Me.pbar_processo.Max = X1
    Me.pbar_processo.Visible = True
    
    For Y = 1 To X1
        X2 = X2 + 1
'        Me.pbar_processo.Value = X2
        Get 11, X2, ArqMovUnimed
        Rem ###################################################################
        Rem abaixo será realizada a consuta da existencia do funcionario na rm
        Rem ###################################################################
'        If Val(Trim(ArqMovUnimed.Fchapa)) = 3466 Then
'           MsgBox ""
'        End If
        
        Set rs = New ADODB.Recordset
        schapa = Format(Val(Trim(ArqMovUnimed.Fchapa)), "000000")
        Set rs = CCTempneUniMvFun.RMFuncionario_Consulta(Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex), schapa)
        If rs.RecordCount = 0 Then
           mfl_grid.Col = 0: mfl_grid.Text = "*"
           mfl_grid.Col = 2: mfl_grid.Text = "Chapa não encontrado"
           mfl_grid.Col = 8: mfl_grid.Text = "Func. não encontrado."
        Else
           mfl_grid.Col = 0: mfl_grid.Text = " "
           mfl_grid.Col = 2: mfl_grid.Text = rs!NOME
           mfl_grid.Col = 8: mfl_grid.Text = " "
        End If
        Rem ###################################################################
        
        Rem ###################################################################
        Rem abaixo sera verificado o mes referenta a importacao
        Rem ###################################################################
''        ddata_Compara = CDate("01/" & Mid$(sAnoMes, 5, 2) & "/" & Mid$(sAnoMes, 1, 4))
''        ddata_Compara = VBA.DateAdd("m", 1, ddata_Compara)
''
''        If Mid$(ddata_Compara, 7, 4) & Mid$(ddata_Compara, 4, 2) <> Mid$(ArqMovUnimed.Fdtpagto, 1, 4) & Mid$(ArqMovUnimed.Fdtpagto, 5, 2) Then
''           mfl_grid.Col = 0: mfl_grid.Text = "*"
''           mfl_grid.Col = 8: mfl_grid.Text = mfl_grid.Text & " # Mes de referencia difere do Movimento."
''        End If
        
        Rem ###################################################################
        Rem abaixo será realizada a consuta da existencia da movimentacao do funcionario
        Rem ###################################################################
        Set rs = New ADODB.Recordset
        Set rs = CCTempneUniMvFun.MovFuncionario_ExisteMov(Me.cbo_coligada.ItemData(Me.cbo_coligada.ListIndex), _
                                                           schapa, _
                                                           sAnoMes)
        If rs.RecordCount = 1 Then
           mfl_grid.Col = 0: mfl_grid.Text = "*"
           mfl_grid.Col = 8: mfl_grid.Text = mfl_grid.Text & " # Já existe Movimento."
        End If
        
        mfl_grid.Col = 1: mfl_grid.Text = schapa
        mfl_grid.Col = 3: mfl_grid.Text = Mid$(ArqMovUnimed.Fdtpagto, 7, 2) & "/" & Mid$(ArqMovUnimed.Fdtpagto, 5, 2) & "/" & Mid$(ArqMovUnimed.Fdtpagto, 1, 4)
        mfl_grid.Col = 4: mfl_grid.Text = ArqMovUnimed.Fhora
        mfl_grid.Col = 5: mfl_grid.Text = Format(VBA.CDbl(ArqMovUnimed.Frefer) / 100, "0.00")
        mfl_grid.Col = 6: mfl_grid.Text = Format(VBA.CDbl(ArqMovUnimed.Fvalor) / 100, "0.00")
        mfl_grid.Col = 7: mfl_grid.Text = ArqMovUnimed.Falterado

        mfl_grid.Rows = mfl_grid.Rows + 1
        mfl_grid.Row = mfl_grid.Row + 1
        
    Next
    
    Close #11
   
End If

Me.mfl_grid.Visible = True
Me.pbar_processo.Visible = False

Exit Sub

Erro:

Call Limpar_mfl_grid
Call Ajuste_Tela
Me.mfl_grid.Visible = True
Me.pbar_processo.Visible = False

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
rs.Fields.Append "DT_EVENTO", ADODB.DataTypeEnum.adVarChar, 10
rs.Fields.Append "HORA", ADODB.DataTypeEnum.adVarChar, 5
rs.Fields.Append "REFERENCIA", ADODB.DataTypeEnum.adDouble
rs.Fields.Append "VALOR", ADODB.DataTypeEnum.adDouble

rs.Open

mfl_grid.Row = 0

If mfl_grid.Rows > 0 Then
   nx = 0
   While nx < Val(Me.txtlidos.Text)
       nx = nx + 1
       mfl_grid.Row = nx
       mfl_grid.Col = 1: rs.AddNew "CHAPA", IIf(IsNull(mfl_grid.Text), " ", Mid$(mfl_grid.Text, 1, 7))
       mfl_grid.Col = 2: rs.Fields("FUNCIONARIO").Value = IIf(IsNull(mfl_grid.Text), " ", Mid$(mfl_grid.Text, 1, 30))
       mfl_grid.Col = 3: rs.Fields("DT_EVENTO").Value = IIf(IsNull(mfl_grid.Text), " ", Mid$(mfl_grid.Text, 1, 10))
       mfl_grid.Col = 4: rs.Fields("HORA").Value = IIf(IsNull(mfl_grid.Text), " ", Mid$(mfl_grid.Text, 2, 5))
       mfl_grid.Col = 5: rs.Fields("REFERENCIA").Value = IIf(IsNull(mfl_grid.Text), " ", CDbl(Trim(mfl_grid.Text)))
       mfl_grid.Col = 6: rs.Fields("VALOR").Value = IIf(IsNull(mfl_grid.Text), " ", CDbl(Trim(mfl_grid.Text)))
       rs.Update
   Wend
Else
   MsgBox "Sem Movimentação, Retorne."
   Exit Sub
End If

Set oTela = New frmRelCristalReport

Me.MousePointer = vbHourglass

Set CrystalReport1 = Application.OpenReport(App.Path & "\crptImportacaoUnimed.rpt")

CrystalReport1.Database.SetDataSource rs

CrystalReport1.ParameterFields(1).AddCurrentValue "teste1 "
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

Private Sub dir_destino_Change()
Me.flie_destino.Path = Me.dir_destino.Path
End Sub

Private Sub dir_destino_Click()
Me.flie_destino.Path = Me.dir_destino.Path
End Sub

Private Sub drv_destino_Change()
Me.dir_destino.Path = Me.drv_destino.Drive
End Sub

Private Sub flie_destino_Click()
If Len(Trim(Me.flie_destino.Path)) > 3 Then
   Me.txt_Arq_Importacao.Text = Me.flie_destino.Path & "\" & Me.flie_destino.FileName
Else
   Me.txt_Arq_Importacao.Text = Me.flie_destino.Path & Me.flie_destino.FileName
End If
End Sub

Private Sub Form_Activate()
If Flag_ativo = True Then
   Exit Sub
End If
Me.Top = 0
Me.Left = 0
Flag_ativo = False
Call carregar_coligada
If Me.cbo_coligada.ListCount = 1 Then Me.cbo_coligada.ListIndex = 0
Flag_ativo = True
Me.cbo_coligada.SetFocus

End Sub

Private Sub Form_Load()
Dim nx As Integer

Me.Top = 0
Me.Left = 0
Call Limpar_mfl_grid
Me.dir_destino.Path = Me.drv_destino.Drive
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
'      Flag_ativo = True
'      Me.cbo_coligada.ListIndex = 0
'      Call Confirmar_Dados_coligada
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
          lbl_Msg_Fechamento.Caption = "Importando Dados Para Periodo de " & Mid$(cRec!TCO_ANO_MES_PROC, 5, 2) & "/" & Mid$(cRec!TCO_ANO_MES_PROC, 1, 4) & ", e desconto de " & Format(cRec!TCO_DESCONTO, "0.00") & " %"
          sPercentual = cRec!TCO_DESCONTO
          sAnoMes = cRec!TCO_ANO_MES_PROC
          If Flag_ativo = True Then
             If cRec!TCO_MOV_ABERTO = 1 Then
                Me.cmd_conf_arquivo.Enabled = False
                Me.cmd_imprime_mov.Visible = False
                Me.cmd_atualizar.Visible = False
                MsgBox "Esta Empresa, está com o movimento fechado, Abra o Mês/Ano para abertura da Importação!"
             Else
                Me.cmd_conf_arquivo.Enabled = True
                Me.cmd_imprime_mov.Visible = True
                Me.cmd_atualizar.Visible = True
             End If
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

Private Sub Ajuste_Tela()

mfl_grid.Col = 0: mfl_grid.ColWidth(0) = 400: mfl_grid.Text = "ST"
mfl_grid.Col = 1: mfl_grid.ColWidth(1) = 1000: mfl_grid.Text = "CHAPA"
mfl_grid.Col = 2: mfl_grid.ColWidth(2) = 2400: mfl_grid.Text = "FUNCIONARIO"
mfl_grid.Col = 3: mfl_grid.ColWidth(3) = 1500:  mfl_grid.Text = "DT.EVENTO"
mfl_grid.Col = 4: mfl_grid.ColWidth(4) = 0: mfl_grid.Text = "HORA"
mfl_grid.Col = 5: mfl_grid.ColWidth(5) = 1000:  mfl_grid.Text = "REFER"
mfl_grid.Col = 6: mfl_grid.ColWidth(6) = 1100:  mfl_grid.Text = "VALOR"
mfl_grid.Col = 7: mfl_grid.ColWidth(7) = 700:  mfl_grid.Text = "ALT"
mfl_grid.Col = 8: mfl_grid.ColWidth(8) = 7000:  mfl_grid.Text = "Observacao"
mfl_grid.Col = 0: mfl_grid.BackColor = &H80FFFF

mfl_grid.Row = 0

mfl_grid.HighLight = False
mfl_grid.ColAlignment(0) = flexAlignCenterCenter
mfl_grid.ColAlignment(1) = flexAlignLeftCenter
mfl_grid.ColAlignment(2) = flexAlignLeftCenter
mfl_grid.ColAlignment(3) = flexAlignCenterCenter
mfl_grid.ColAlignment(4) = flexAlignCenterCenter
mfl_grid.ColAlignment(5) = flexAlignRightCenter
mfl_grid.ColAlignment(6) = flexAlignRightCenter
mfl_grid.ColAlignment(7) = flexAlignCenterCenter

End Sub

