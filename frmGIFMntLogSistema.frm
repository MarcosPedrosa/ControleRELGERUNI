VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGIFMntLogSistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log do Sistema"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   10755
   Begin VB.Frame Frame3 
      Caption         =   "Intervalo / Usuário"
      Height          =   1605
      Left            =   150
      TabIndex        =   4
      Top             =   120
      Width           =   10515
      Begin VB.CommandButton cmd_Pesquisar 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Pesquisar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   630
         Width           =   1245
      End
      Begin VB.ComboBox cbo_usuario 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Selecione o usuário para seleção."
         Top             =   720
         Width           =   3465
      End
      Begin VB.ComboBox cbo_acao 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Selecione o usuário para seleção."
         Top             =   300
         Width           =   3465
      End
      Begin VB.CheckBox CHK_TODOS_PERIODOS 
         Caption         =   "Todos os periodos"
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
         Left            =   4890
         TabIndex        =   8
         Top             =   1170
         Width           =   1965
      End
      Begin VB.Frame Frame1 
         Caption         =   "Classifiicação relatório"
         Height          =   825
         Left            =   4890
         TabIndex        =   5
         Top             =   210
         Width           =   2535
         Begin VB.OptionButton Opt_Usuario 
            Caption         =   "Usuário"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "cronológico"
            Height          =   255
            Left            =   1140
            TabIndex        =   6
            Top             =   360
            Width           =   1155
         End
      End
      Begin MSComCtl2.DTPicker dt_inicio 
         Height          =   315
         Left            =   1290
         TabIndex        =   12
         Top             =   1140
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66977793
         CurrentDate     =   40148
      End
      Begin MSComCtl2.DTPicker dt_final 
         Height          =   315
         Left            =   3330
         TabIndex        =   13
         Top             =   1140
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   66977793
         CurrentDate     =   40148
      End
      Begin VB.Label Label3 
         Caption         =   "Periodo.:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1140
         Width           =   1125
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
         Left            =   2820
         TabIndex        =   16
         Top             =   1140
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "Usuário.:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   727
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ação.:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   780
      End
   End
   Begin VB.CommandButton cmd_imprime_pallet 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   345
      Left            =   8040
      Picture         =   "frmGIFMntLogSistema.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Imprimir os Pallet's encontrados"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   9420
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1275
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   0
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label Label4 
      Caption         =   "Total registros : "
      Height          =   225
      Left            =   210
      TabIndex        =   2
      Top             =   1860
      Width           =   1185
   End
End
Attribute VB_Name = "frmGIFMntLogSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Private Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Private cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public Confirma_Mudanca As Boolean 'Servirá para confirmar as mudanças de alteracoes dos campos na tela


Private Sub cmd_imprime_pallet_Click()
Dim oTela As frmRelCristalReport
Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim rs As ADODB.Recordset
Dim nx As Integer
Dim nValor As Double

On Error GoTo Erro

Set rs = New ADODB.Recordset

rs.Fields.Append "LOG_DATA", ADODB.DataTypeEnum.adDBDate
rs.Fields.Append "LOG_HORA", ADODB.DataTypeEnum.adVarChar, 8
rs.Fields.Append "LOG_USU", ADODB.DataTypeEnum.adVarChar, 10
rs.Fields.Append "LOG_TABELA", ADODB.DataTypeEnum.adVarChar, 30
rs.Fields.Append "LOG_ACAO", ADODB.DataTypeEnum.adVarChar, 50
rs.Fields.Append "LOG_SQL", ADODB.DataTypeEnum.adVarChar, 500
rs.Fields.Append "LOG_OBSERVACAO", ADODB.DataTypeEnum.adVarChar, 200

rs.Open

cRec.MoveFirst

If cRec.RecordCount > 0 Then
'   Me.txtlidos.Text = cRec.RecordCount
   cRec.MoveFirst
   While Not cRec.EOF
       rs.AddNew "LOG_DATA", cRec!LOG_DATA
       rs.Fields("LOG_HORA").Value = cRec!LOG_HORA
       rs.Fields("LOG_USU").Value = cRec!LOG_USU
       rs.Fields("LOG_TABELA").Value = IIf(Len(Trim((cRec!LOG_TABELA))) = " ", 0, cRec!LOG_TABELA)
       rs.Fields("LOG_ACAO").Value = IIf(Len(Trim((cRec!LOG_ACAO))) = " ", 0, cRec!LOG_ACAO)
       rs.Fields("LOG_SQL").Value = IIf(Len(Trim((cRec!LOG_SQL))) = " ", 0, cRec!LOG_SQL)
       rs.Fields("LOG_OBSERVACAO").Value = IIf(Len(Trim((cRec!LOG_OBSERVACAO))) = 0, 0, cRec!LOG_OBSERVACAO)
       rs.Update
       cRec.MoveNext
   Wend
Else
   MsgBox "Sem movimentação, Retorne."
   Exit Sub
End If

Set oTela = New frmRelCristalReport

Me.MousePointer = vbHourglass

Set CrystalReport1 = Application.OpenReport(App.Path & "\crptAcaoUsuario.rpt")

CrystalReport1.Database.SetDataSource rs

CrystalReport1.ParameterFields(1).AddCurrentValue "no periodo de " & Me.dt_inicio.Value & " a " & Me.dt_final.Value
CrystalReport1.ParameterFields(1).DiscreteOrRangeKind = crDiscreteValue

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
Dim sClassf As String
Dim sAcao As String
Dim sUsur As String

On Error GoTo Erro

'
'Call Limpar_mfl_grid
'
If Me.CHK_TODOS_PERIODOS.Value = 0 Then
   If Me.dt_inicio.Value > Me.dt_final.Value Then
      MsgBox "Data dos pallets, o inicio está maior que o final, redigite!"
      Me.dt_inicio.SetFocus
      Exit Sub
   Else
      sDataINI = Format(Me.dt_inicio.Value, "yyyymmdd")
      sDataFIM = Format(Me.dt_final.Value, "yyyymmdd")
   End If
Else
      sDataINI = ""
      sDataFIM = ""
End If

If Me.Opt_Usuario.Value = True Then
   sClassf = 1
Else
   sClassf = 2
End If

If Me.cbo_acao.ListIndex = 0 Then
   sAcao = ""
Else
   sAcao = Me.cbo_acao.List(Me.cbo_acao.ListIndex)
End If

If Me.cbo_usuario.ListIndex = 0 Then
   sUsur = ""
Else
   sUsur = Me.cbo_usuario.List(Me.cbo_usuario.ListIndex)
End If

Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass

Set cRec = CCTemp.LOG_CONSULTAR(sNomeBanco, _
                                sDataINI, _
                                sDataFIM, _
                                sClassf, _
                                sAcao, _
                                sUsur)


If cRec.RecordCount > 0 Then
   Me.cmd_imprime_pallet.Enabled = True
Else
   Me.cmd_imprime_pallet.Enabled = False
End If

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
Me.dt_inicio.Value = Format(Now(), "dd/mm/yyyy")
Me.dt_final.Value = Format(Now(), "dd/mm/yyyy")

Call Carrega_Acoes
Call Carrega_Acoes_Usuario

End Sub
Function Carrega_Acoes()
Dim nx As Integer

On Error GoTo Erro

Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass
Set cRec = CCTemp.LOG_CONSULTAR_Acoes(sNomeBanco)

Me.cbo_acao.Clear
Me.cbo_acao.AddItem "TODOS AS AÇÕES"
Me.cbo_acao.ItemData(0) = 0

If cRec.RecordCount > 0 Then
   cRec.MoveFirst
   While Not cRec.EOF
       If Not IsNull(cRec!LOG_ACAO) Then
          nx = nx + 1
          Me.cbo_acao.AddItem Trim(cRec!LOG_ACAO)
       End If
       cRec.MoveNext
   Wend
   Me.cbo_acao.ListIndex = 0
Else
   MsgBox "Não existem Ações realizadas, procure o responsável."
End If

Me.MousePointer = vbDefault

Rem Set cRec = Nothing
Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Function
Function Carrega_Acoes_Usuario()
Dim nx As Integer

On Error GoTo Erro

Set cRec = New ADODB.Recordset

Me.MousePointer = vbHourglass
Set cRec = CCTemp.LOG_CONSULTAR_Acoes_Usuario(sNomeBanco)

Me.cbo_usuario.Clear
Me.cbo_usuario.AddItem "TODOS OS USUÁRIOS"
Me.cbo_usuario.ItemData(0) = 0

If cRec.RecordCount > 0 Then
   cRec.MoveFirst
   While Not cRec.EOF
       If Not IsNull(cRec!LOG_USU) Then
          nx = nx + 1
          Me.cbo_usuario.AddItem cRec!LOG_USU
       End If
       cRec.MoveNext
   Wend
Else
   MsgBox "Não existem Usuários com ações realizadas, procure o responsável."
End If

Me.cbo_usuario.ListIndex = 0
Me.MousePointer = vbDefault

Rem Set cRec = Nothing
Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault
End Function

Public Function CCTemp() As neLog
     Set CCTemp = New neLog
End Function



