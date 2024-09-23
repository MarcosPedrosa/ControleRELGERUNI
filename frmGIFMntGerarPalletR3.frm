VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Begin VB.Form frmGIFMntGerarPalletR3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerar pallet's para o R3"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   11505
   Begin VB.TextBox Txt_arquivo 
      Height          =   315
      Left            =   4920
      TabIndex        =   20
      Text            =   "\Arq_Pallet_R3.txt"
      Top             =   5280
      Width           =   3495
   End
   Begin VB.DriveListBox Drv_Destino 
      Height          =   315
      Left            =   3870
      TabIndex        =   19
      Top             =   5280
      Width           =   1035
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   1230
      MaxLength       =   6
      TabIndex        =   16
      Top             =   5280
      Width           =   1005
   End
   Begin VB.CommandButton cmd_Gerar 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Gerar arquivo"
      Enabled         =   0   'False
      Height          =   330
      Left            =   8430
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5280
      Width           =   1275
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   10140
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5280
      Width           =   1275
   End
   Begin vsOcx6LibCtl.vsElastic vsElastic1 
      Height          =   5130
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   9049
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      Appearance      =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      _GridInfo       =   ""
      Begin VB.Frame Frame2 
         Height          =   3675
         Left            =   150
         TabIndex        =   12
         Top             =   1350
         Width           =   10545
         Begin MSFlexGridLib.MSFlexGrid mfl_gridcomp 
            Height          =   3345
            Left            =   90
            TabIndex        =   13
            Top             =   180
            Width           =   10395
            _ExtentX        =   18336
            _ExtentY        =   5900
            _Version        =   393216
            Cols            =   9
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
      End
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   150
         TabIndex        =   2
         Top             =   30
         Width           =   10515
         Begin VB.Frame Frame1 
            Caption         =   "Intervalo de palet's"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   180
            TabIndex        =   4
            Top             =   180
            Width           =   7935
            Begin VB.ComboBox CBO_SEQ1 
               Height          =   315
               Left            =   2280
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   330
               Width           =   780
            End
            Begin VB.ComboBox CBO_SEQ2 
               Height          =   315
               Left            =   5070
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Top             =   330
               Width           =   780
            End
            Begin VB.CheckBox chk_pallet 
               Caption         =   "Todos pallets"
               Height          =   255
               Left            =   6450
               TabIndex        =   5
               Top             =   360
               Width           =   1275
            End
            Begin MSComCtl2.DTPicker dt_inicio 
               Height          =   315
               Left            =   870
               TabIndex        =   8
               Top             =   330
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               _Version        =   393216
               Format          =   67371009
               CurrentDate     =   40148
            End
            Begin MSComCtl2.DTPicker dt_final 
               Height          =   315
               Left            =   3660
               TabIndex        =   9
               Top             =   330
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               _Version        =   393216
               Format          =   67371009
               CurrentDate     =   40148
            End
            Begin VB.Label Label3 
               Caption         =   "De.:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   210
               TabIndex        =   11
               Top             =   330
               Width           =   465
            End
            Begin VB.Label Label2 
               Caption         =   "até"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3120
               TabIndex        =   10
               Top             =   330
               Width           =   435
            End
         End
         Begin VB.CommandButton cmd_Pesquisar 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Pesquisar"
            Height          =   375
            Left            =   8850
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmd_imprime_pallet 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   405
         Left            =   10740
         Picture         =   "frmGIFMntGerarPalletR3.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimir os Pallet's encontrados"
         Top             =   1440
         Width           =   525
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Nome do Arquivo : "
      Height          =   225
      Left            =   2430
      TabIndex        =   18
      Top             =   5340
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Total registros : "
      Height          =   225
      Left            =   0
      TabIndex        =   17
      Top             =   5340
      Width           =   1185
   End
End
Attribute VB_Name = "frmGIFMntGerarPalletR3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Private Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Private cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public Confirma_Mudanca As Boolean 'Servirá para confirmar as mudanças de alteracoes dos campos na tela

Private Sub cmd_Gerar_Click()
Dim x As Double

On Error GoTo Erro

Me.MousePointer = vbHourglass

Open Me.Drv_Destino.Drive & Txt_arquivo.Text For Random Access Read Write As #1 Len = Len(ArqR3)
Close 1
Kill Txt_arquivo.Text
Open Txt_arquivo.Text For Random Access Read Write As #1 Len = Len(ArqR3)
cRec.MoveFirst

If cRec.RecordCount > 0 Then
'   Me.txtlidos.Text = cRec.RecordCount
   cRec.MoveFirst
   x = 0
   While Not cRec.EOF
        ArqR3.Fetiqueta = cRec!NUM_CAIXA
        ArqR3.Fpallet = cRec!PALLET
        ArqR3.FFinal = Chr$(13) + Chr$(10)
        x = x + 1
        Put 1, x, ArqR3
       cRec.MoveNext
   Wend
   Close #1
   Call CCTemp.MANUTENCAO_INVENTARIO_GeraPaletR3(sNomeBanco, sNomeUsuario, Format(Now(), "dd/mm/yyyy"), Format(cRec.RecordCount, "0"))
   MsgBox "Arquivo gerado com sucesso."
   
Else
   Close #1
   MsgBox "Sem movimentação, Retorne."
End If

Me.MousePointer = vbDefault
Exit Sub

Erro:
Close #1
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault
End Sub

Private Sub cmd_imprime_pallet_Click()
Dim oTela As frmRelCristalReport
Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim rs As ADODB.Recordset
Dim nx As Integer
Dim nValor As Double

On Error GoTo Erro

Set rs = New ADODB.Recordset

rs.Fields.Append "NUM_CAIXA", ADODB.DataTypeEnum.adVarChar, 10
rs.Fields.Append "TIPO_CAIXA", ADODB.DataTypeEnum.adVarChar, 8
rs.Fields.Append "PECA", ADODB.DataTypeEnum.adVarChar, 15
rs.Fields.Append "LOTE", ADODB.DataTypeEnum.adVarChar, 15
rs.Fields.Append "NF_VENDA", ADODB.DataTypeEnum.adVarChar, 16
rs.Fields.Append "ORDEM_VENDA", ADODB.DataTypeEnum.adVarChar, 11
rs.Fields.Append "SEQUENCIA", ADODB.DataTypeEnum.adDouble
rs.Fields.Append "PLACA", ADODB.DataTypeEnum.adVarChar, 15
rs.Fields.Append "QTDEITENS", ADODB.DataTypeEnum.adDouble

rs.Open

cRec.MoveFirst

If cRec.RecordCount > 0 Then
'   Me.txtlidos.Text = cRec.RecordCount
   cRec.MoveFirst
   While Not cRec.EOF
       rs.AddNew "NUM_CAIXA", cRec!NUM_CAIXA
       rs.Fields("TIPO_CAIXA").Value = cRec!TIPO_CAIXA
       rs.Fields("PECA").Value = cRec!PECA
       rs.Fields("LOTE").Value = IIf(Len(Trim((cRec!LOTE))) = 0, 0, cRec!LOTE)
       rs.Fields("NF_VENDA").Value = IIf(Len(Trim((cRec!NF_VENDA))) = 0, 0, cRec!NF_VENDA)
       rs.Fields("ORDEM_VENDA").Value = IIf(Len(Trim((cRec!ORDEM_VENDA))) = 0, 0, cRec!ORDEM_VENDA)
       rs.Fields("SEQUENCIA").Value = IIf(Len(Trim((cRec!SEQUENCIA))) = 0, 0, cRec!SEQUENCIA)
       rs.Fields("PLACA").Value = IIf(Len(Trim((cRec!PALLET))) = 0, 0, cRec!PALLET)
       rs.Fields("QTDEITENS").Value = IIf(Len(Trim((cRec!QTDE_NA_CAIXA))) = 0, 0, cRec!QTDE_NA_CAIXA)
       rs.Update
       cRec.MoveNext
   Wend
Else
   MsgBox "Sem movimentação, Retorne."
   Exit Sub
End If

Set oTela = New frmRelCristalReport

Me.MousePointer = vbHourglass

Set CrystalReport1 = Application.OpenReport(App.Path & "\crptConferenciaR3.rpt")

CrystalReport1.Database.SetDataSource rs

CrystalReport1.ParameterFields(1).AddCurrentValue "444"
CrystalReport1.ParameterFields(1).DiscreteOrRangeKind = crDiscreteValue

CrystalReport1.ParameterFields(2).AddCurrentValue "555 "
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

Private Sub cmd_Pesquisar_Click()
Dim nx As Integer
Dim sDataINI As String
Dim sDataFIM As String
Dim sSEQ1 As String
Dim sSEQ2 As String
Dim sStatus As String
Dim sCliente As String


On Error GoTo Erro

'
'Call Limpar_mfl_grid
'

If Me.chk_pallet.Value = 1 Then
   sDataINI = ""
   sDataFIM = ""
   sSEQ1 = ""
   sSEQ2 = ""
Else
   If Me.dt_inicio.Value > Me.dt_final.Value Then
      MsgBox "Data dos pallets, o inicio está maior que o final, redigite!"
      Me.dt_inicio.SetFocus
      Exit Sub
   Else
      sDataINI = Format(Me.dt_inicio.Value, "yyyymmdd")
      sDataFIM = Format(Me.dt_final.Value, "yyyymmdd")
   End If
   sSEQ1 = Me.CBO_SEQ1.List(Me.CBO_SEQ1.ListIndex)
   sSEQ2 = Me.CBO_SEQ2.List(Me.CBO_SEQ2.ListIndex)
End If
   

Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass

Set cRec = CCTemp.MANUTENCAO_Consultar_Mov_Para_R3(sNomeBanco, _
                                                   sDataINI, _
                                                   sDataFIM, _
                                                   sSEQ1, _
                                                   sSEQ2)


Me.mfl_gridcomp.Visible = False
mfl_gridcomp.Row = 0
mfl_gridcomp.Col = 0: mfl_gridcomp.ColWidth(0) = 1300: mfl_gridcomp.Text = "NºCAIXA"
mfl_gridcomp.Col = 1: mfl_gridcomp.ColWidth(1) = 1200: mfl_gridcomp.Text = "CAIXA"
mfl_gridcomp.Col = 2: mfl_gridcomp.ColWidth(2) = 1000: mfl_gridcomp.Text = "PECA"
mfl_gridcomp.Col = 3: mfl_gridcomp.ColWidth(3) = 1200: mfl_gridcomp.Text = "LOTE"
mfl_gridcomp.Col = 4: mfl_gridcomp.ColWidth(4) = 600: mfl_gridcomp.Text = "QTD"
mfl_gridcomp.Col = 5: mfl_gridcomp.ColWidth(5) = 1100: mfl_gridcomp.Text = "N.FISCAL"
mfl_gridcomp.Col = 6: mfl_gridcomp.ColWidth(6) = 1350: mfl_gridcomp.Text = "ORD.VENDA"
mfl_gridcomp.Col = 7: mfl_gridcomp.ColWidth(7) = 900: mfl_gridcomp.Text = "SEQ."
mfl_gridcomp.Col = 8: mfl_gridcomp.ColWidth(8) = 1400: mfl_gridcomp.Text = "PALLET"
'mfl_gridcomp.Col = 2: mfl_gridcomp.BackColor = &H80FFFF

mfl_gridcomp.Row = 0

mfl_gridcomp.HighLight = False
mfl_gridcomp.ColAlignment(0) = flexAlignCenterCenter
mfl_gridcomp.ColAlignment(1) = flexAlignCenterCenter
mfl_gridcomp.ColAlignment(2) = flexAlignLeftCenter
mfl_gridcomp.ColAlignment(3) = flexAlignLeftCenter
mfl_gridcomp.ColAlignment(4) = flexAlignLeftCenter
mfl_gridcomp.Row = 1

If cRec.RecordCount > 0 Then
   Me.txtlidos.Text = cRec.RecordCount
   Me.cmd_Gerar.Enabled = True
   Me.cmd_imprime_pallet.Enabled = True
   cRec.MoveFirst
   For nx = 1 To cRec.RecordCount
       mfl_gridcomp.Col = 0: mfl_gridcomp.Text = cRec.Fields("NUM_CAIXA")
       mfl_gridcomp.Col = 1: mfl_gridcomp.Text = cRec.Fields("TIPO_CAIXA")
       mfl_gridcomp.Col = 2: mfl_gridcomp.Text = cRec.Fields("PECA")
       mfl_gridcomp.Col = 3: mfl_gridcomp.Text = cRec.Fields("LOTE")
       mfl_gridcomp.Col = 4: mfl_gridcomp.Text = cRec.Fields("QTDE_NA_CAIXA")
       mfl_gridcomp.Col = 5: mfl_gridcomp.Text = cRec.Fields("NF_VENDA")
       mfl_gridcomp.Col = 6: mfl_gridcomp.Text = cRec.Fields("ORDEM_VENDA")
       mfl_gridcomp.Col = 7: mfl_gridcomp.Text = cRec.Fields("SEQUENCIA")
       mfl_gridcomp.Col = 8: mfl_gridcomp.Text = cRec.Fields("PALLET")
       cRec.MoveNext
       If Not cRec.EOF Then
          mfl_gridcomp.Rows = mfl_gridcomp.Rows + 1
          mfl_gridcomp.Row = mfl_gridcomp.Row + 1
       End If
   Next
Else
   Me.cmd_Gerar.Enabled = False
   Me.cmd_imprime_pallet.Enabled = False
End If

Me.mfl_gridcomp.Visible = True

Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description
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
'   Else
'       SendKeys "+{TAB}" ' retornar campo
'   End If
End If


End Sub

Private Sub Form_Load()
Dim nx As Integer

Me.Top = 0
Me.Left = 0

For nx = 1 To 99
    Me.CBO_SEQ1.AddItem nx
    Me.CBO_SEQ2.AddItem nx
Next


Me.CBO_SEQ1.ListIndex = 0
Me.CBO_SEQ2.ListIndex = 0
Me.dt_inicio.Value = Format(Now(), "dd/mm/yyyy")
Me.dt_final.Value = Format(Now(), "dd/mm/yyyy")

mfl_gridcomp.Row = 0
mfl_gridcomp.Col = 0: mfl_gridcomp.ColWidth(0) = 1500: mfl_gridcomp.Text = "NºCAIXA"
mfl_gridcomp.Col = 1: mfl_gridcomp.ColWidth(1) = 1200: mfl_gridcomp.Text = "CAIXA"
mfl_gridcomp.Col = 2: mfl_gridcomp.ColWidth(2) = 1200: mfl_gridcomp.Text = "PECA"
mfl_gridcomp.Col = 3: mfl_gridcomp.ColWidth(3) = 1200: mfl_gridcomp.Text = "LOTE"
mfl_gridcomp.Col = 4: mfl_gridcomp.ColWidth(4) = 700: mfl_gridcomp.Text = "QTDE"
mfl_gridcomp.Col = 5: mfl_gridcomp.ColWidth(5) = 1250: mfl_gridcomp.Text = "N.FISCAL"
mfl_gridcomp.Col = 6: mfl_gridcomp.ColWidth(6) = 1300: mfl_gridcomp.Text = "ORD.VENDA"
mfl_gridcomp.Col = 7: mfl_gridcomp.ColWidth(7) = 1400: mfl_gridcomp.Text = "SEQUENCIA"
mfl_gridcomp.Col = 8: mfl_gridcomp.ColWidth(8) = 1100: mfl_gridcomp.Text = "PALLET"

End Sub

Public Function CCTemp() As neManutencao
     Set CCTemp = New neManutencao
End Function

