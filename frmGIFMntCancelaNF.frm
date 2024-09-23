VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGIFMntCancelaNF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelamento de Nota Fiscal"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   9930
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5910
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   5745
      Left            =   150
      TabIndex        =   4
      Top             =   90
      Width           =   9675
      Begin VB.TextBox txtlidos 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H8000000F&
         Height          =   315
         Left            =   1530
         MaxLength       =   6
         TabIndex        =   9
         Top             =   5220
         Width           =   1005
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ítens"
         Height          =   3585
         Left            =   240
         TabIndex        =   8
         Top             =   1380
         Width           =   9255
         Begin MSFlexGridLib.MSFlexGrid mfl_gridcomp 
            Height          =   3135
            Left            =   90
            TabIndex        =   2
            Top             =   270
            Width           =   9045
            _ExtentX        =   15954
            _ExtentY        =   5530
            _Version        =   393216
            Cols            =   7
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
      Begin VB.CommandButton cmd_cancelamento 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   555
         Left            =   8700
         Picture         =   "frmGIFMntCancelaNF.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cancelar Nota Fiscal"
         Top             =   5100
         Width           =   795
      End
      Begin VB.Frame Frame2 
         Caption         =   "Selecionar Nota fiscal"
         Height          =   1005
         Left            =   270
         TabIndex        =   5
         Top             =   270
         Width           =   9225
         Begin VB.CommandButton cmd_confirmar_NF 
            BackColor       =   &H00FF8080&
            Height          =   375
            Left            =   6060
            Picture         =   "frmGIFMntCancelaNF.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Localizar Nota fiscal"
            Top             =   420
            Width           =   435
         End
         Begin VB.TextBox TXT_NF 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3720
            MaxLength       =   10
            TabIndex        =   0
            ToolTipText     =   "Digite o Numero da Nota fiscal e tecle no botao ao lado "
            Top             =   390
            Width           =   2805
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Nº Nota Fiscal.:"
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
            Left            =   1830
            TabIndex        =   6
            Top             =   420
            Width           =   1875
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Total registros : "
         Height          =   285
         Left            =   330
         TabIndex        =   10
         Top             =   5250
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmGIFMntCancelaNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Private Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Private cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public Confirma_Mudanca As Boolean 'Servirá para confirmar as mudanças de alteracoes dos campos na tela

Private Sub cmd_cancelamento_Click()
Dim nx As Integer
Dim sNF As String
Dim nRet As Integer

On Error GoTo Erro

Set cRec = New ADODB.Recordset

If Len(Trim(Me.TXT_NF.Text)) = 0 Then
   MsgBox "Digite o número de uma Nota fiscal para ser validada!"
   Me.MousePointer = vbDefault
   Me.TXT_NF.Text = ""
   Me.TXT_NF.SetFocus
   Exit Sub
Else
   sNF = Format(Trim(Me.TXT_NF.Text), "0")
End If

nRet = MsgBox("Confirma cancelamento?", vbQuestion & vbYesNo, Me.Caption)
'Se confirmou a exclusão:
If nRet = 6 Then
   Me.MousePointer = vbHourglass
   Set cRec = CCTemp.MANUTENCAO_NF_Cancelar(sNomeBanco, sNF)
   Call Limpar_mfl_gridcompcomp
   MsgBox "Nota fiscal cancelada com Successo!"
   Me.cmd_cancelamento.Enabled = False
   Me.TXT_NF.Text = ""
   Me.txtlidos.Text = ""
End If

Me.TXT_NF.SetFocus

Set cRec = Nothing
Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub

Private Sub cmd_confirmar_NF_Click()
Call Limpar_mfl_gridcompcomp
Call Confirmar_NF
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
'Call Limpar_campos
'Call Desabilitar_Campos

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
'ElseIf Me.ActiveControl.TabIndex = 50 Then
'       If KeyCode = 35 Then
'          SendKeys "{END}" ' FIM DO campo
'       End If
End If


End Sub

Private Sub Form_Load()
Dim nx As Integer

Me.Top = 0
Me.Left = 0

Call Limpar_mfl_gridcompcomp

End Sub

Public Function CCTemp() As neManutencao
     Set CCTemp = New neManutencao
End Function

Private Sub Limpar_mfl_gridcompcomp()
Dim nx As Double
Dim nLinhas As Double
Dim nLinhas1 As Double

Me.mfl_gridcomp.Visible = False
mfl_gridcomp.Clear
nLinhas = mfl_gridcomp.Rows

If mfl_gridcomp.Rows > 2 Then
   For nx = mfl_gridcomp.Rows To nLinhas1 - 2 Step -1
       If nx > 2 Then mfl_gridcomp.RemoveItem (nx)
   Next
End If

mfl_gridcomp.Row = 0
mfl_gridcomp.Col = 0: mfl_gridcomp.ColWidth(0) = 1500: mfl_gridcomp.Text = "NºCAIXA"
mfl_gridcomp.Col = 1: mfl_gridcomp.ColWidth(1) = 1200: mfl_gridcomp.Text = "PECA"
mfl_gridcomp.Col = 2: mfl_gridcomp.ColWidth(2) = 1000: mfl_gridcomp.Text = "QTDE"
mfl_gridcomp.Col = 3: mfl_gridcomp.ColWidth(3) = 1300: mfl_gridcomp.Text = "LOTE"
mfl_gridcomp.Col = 4: mfl_gridcomp.ColWidth(4) = 1200: mfl_gridcomp.Text = "TP.CAIXA"
mfl_gridcomp.Col = 5: mfl_gridcomp.ColWidth(5) = 1300: mfl_gridcomp.Text = "PALLET"
mfl_gridcomp.Col = 6: mfl_gridcomp.ColWidth(6) = 1100: mfl_gridcomp.Text = "PLACA"
'mfl_gridcomp.Col = 2: mfl_gridcomp.BackColor = &H80FFFF

mfl_gridcomp.Row = 0

mfl_gridcomp.HighLight = False
mfl_gridcomp.ColAlignment(0) = flexAlignCenterCenter
mfl_gridcomp.ColAlignment(1) = flexAlignCenterCenter
mfl_gridcomp.ColAlignment(2) = flexAlignRightCenter
mfl_gridcomp.ColAlignment(3) = flexAlignLeftCenter
mfl_gridcomp.ColAlignment(4) = flexAlignLeftCenter

Me.mfl_gridcomp.Visible = True

End Sub

Function Confirmar_NF()
Dim nx As Integer
Dim sNF As String

On Error GoTo Erro

Set cRec = New ADODB.Recordset

If Len(Trim(Me.TXT_NF.Text)) = 0 Then
   MsgBox "Digite o número de uma Nota fiscal para ser validada!"
   Me.MousePointer = vbDefault
   Me.TXT_NF.Text = ""
   Me.TXT_NF.SetFocus
   Exit Function
Else
   sNF = Format(Trim(Me.TXT_NF.Text), "0")
End If

Me.MousePointer = vbHourglass

Set cRec = CCTemp.MANUTENCAO_NF_Consultar(sNomeBanco, sNF)

Me.mfl_gridcomp.Visible = False
mfl_gridcomp.Row = 0
mfl_gridcomp.Col = 0: mfl_gridcomp.ColWidth(0) = 1500: mfl_gridcomp.Text = "NºCAIXA"
mfl_gridcomp.Col = 1: mfl_gridcomp.ColWidth(1) = 1200: mfl_gridcomp.Text = "PECA"
mfl_gridcomp.Col = 2: mfl_gridcomp.ColWidth(2) = 1000: mfl_gridcomp.Text = "QTDE"
mfl_gridcomp.Col = 3: mfl_gridcomp.ColWidth(3) = 1300: mfl_gridcomp.Text = "LOTE"
mfl_gridcomp.Col = 4: mfl_gridcomp.ColWidth(4) = 1200: mfl_gridcomp.Text = "TP.CAIXA"
mfl_gridcomp.Col = 5: mfl_gridcomp.ColWidth(5) = 1300: mfl_gridcomp.Text = "PALLET"
mfl_gridcomp.Col = 6: mfl_gridcomp.ColWidth(6) = 1100: mfl_gridcomp.Text = "PLACA"
'mfl_gridcomp.Col = 2: mfl_gridcomp.BackColor = &H80FFFF

mfl_gridcomp.Row = 0

mfl_gridcomp.HighLight = False
mfl_gridcomp.ColAlignment(0) = flexAlignCenterCenter
mfl_gridcomp.ColAlignment(1) = flexAlignCenterCenter
mfl_gridcomp.ColAlignment(2) = flexAlignRightCenter
mfl_gridcomp.ColAlignment(3) = flexAlignLeftCenter
mfl_gridcomp.ColAlignment(4) = flexAlignLeftCenter
mfl_gridcomp.Row = 1

If cRec.RecordCount > 0 Then
   Me.cmd_cancelamento.Enabled = True
   Me.txtlidos.Text = cRec.RecordCount
   cRec.MoveFirst
   For nx = 1 To cRec.RecordCount
       mfl_gridcomp.Col = 0: mfl_gridcomp.Text = cRec.Fields("ID_ETIQUETA")
       mfl_gridcomp.Col = 1: mfl_gridcomp.Text = cRec.Fields("ID_PECA")
       mfl_gridcomp.Col = 2: mfl_gridcomp.Text = cRec.Fields("QTDE")
       mfl_gridcomp.Col = 3: mfl_gridcomp.Text = cRec.Fields("LOTE")
       mfl_gridcomp.Col = 4: mfl_gridcomp.Text = cRec.Fields("TIPO_CAIXA")
       mfl_gridcomp.Col = 5: mfl_gridcomp.Text = cRec.Fields("PALLET")
       mfl_gridcomp.Col = 6: mfl_gridcomp.Text = cRec.Fields("PLACA")
       cRec.MoveNext
       If Not cRec.EOF Then
          mfl_gridcomp.Rows = mfl_gridcomp.Rows + 1
          mfl_gridcomp.Row = mfl_gridcomp.Row + 1
       End If
   Next
Else
   MsgBox "Não existe nota fiscal com esta numeração, redigite!"
   Me.cmd_cancelamento.Enabled = False
   Me.TXT_NF.Text = ""
   Me.TXT_NF.SetFocus
End If

Me.mfl_gridcomp.Visible = True

Set cRec = Nothing
Me.MousePointer = vbDefault

Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault

End Function
Private Sub TXT_NF_Change()
Me.cmd_cancelamento.Enabled = False

If Not Testa_Numerico(Me.TXT_NF.Text, Len(Me.TXT_NF.Text)) Then
   MsgBox "Só aceita numeros, redigite"
   Me.TXT_NF.Text = Mid$(Me.TXT_NF.Text, 1, Len(Trim(Me.TXT_NF.Text)) - 1)
   Me.TXT_NF.SetFocus
   SendKeys "{END}"
End If

End Sub

Private Sub TXT_NF_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Len(Trim(Me.TXT_NF.Text)) > 0 Then
   Call Limpar_mfl_gridcompcomp
   Call Confirmar_NF
End If
End Sub
