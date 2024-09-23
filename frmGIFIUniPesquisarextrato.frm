VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGIFIUniPesquisarextrato 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extrato Funcionário"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3765
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   7725
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3375
         Left            =   150
         TabIndex        =   3
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5953
         _Version        =   393216
         Cols            =   5
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
         TabIndex        =   4
         Top             =   1410
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdfechar 
      Caption         =   "&Fechar"
      Height          =   360
      Left            =   6510
      TabIndex        =   1
      Top             =   4410
      Width           =   1275
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   345
      Left            =   1230
      MaxLength       =   6
      TabIndex        =   0
      Top             =   4410
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   1470
      TabIndex        =   6
      Top             =   90
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Total registros : "
      Height          =   315
      Left            =   30
      TabIndex        =   5
      Top             =   4440
      Width           =   1185
   End
End
Attribute VB_Name = "frmGIFIUniPesquisarextrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ccodigo_pesquisa As String 'Codigo escolhido pelo usuário.
Public cnome As String 'Nome do escolhido pelo usuário.
Public cColigada As String 'EMPRESA COLIGADA
Public rs As ADODB.Recordset
Public nTeclou_Enter As Integer

Private Sub cmdfechar_Click()
ccodigo_pesquisa = ""
cnome = ""
Me.Hide
End Sub


Private Sub Form_Activate()
'   txtNome.SetFocus
'   Me.Grid1.SetFocus
   nTeclou_Enter = 0
End Sub

Private Sub Form_Load()
   Call Carrega_januspesquisa
End Sub
Function Carrega_januspesquisa()

Dim nx As Double
Dim nLinhas As Double

On Error GoTo Erro

Me.Grid1.Visible = False
Me.MousePointer = vbHourglass
Set rs = New ADODB.Recordset

Rem VERIFICAR OS FUNCIONARIOS QUE EXISTEM SALDO NO CADASTRO

Set rs = CCTempneUniMvFun.MovFuncionario_ConsultaExtrato(cColigada, ccodigo_pesquisa)

If rs.RecordCount > 0 Then
   Call Limpar_Grid
   Call Carregar_Grid
End If

Me.MousePointer = vbDefault
Me.Grid1.Visible = True

Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault
   
End Function

Private Sub Limpar_Grid()
Dim nx As Double
Dim nLinhas As Double
Dim nLinhas1 As Double

Grid1.Clear
nLinhas = Grid1.Rows

If Grid1.Rows > 2 Then
   For nx = Grid1.Rows To nLinhas1 - 2 Step -1
       If nx > 2 Then Grid1.RemoveItem (nx)
   Next
End If

Grid1.Row = 0
Grid1.Col = 0: Grid1.ColWidth(0) = 1200:  Grid1.Text = "ANO/MES"
Grid1.Col = 1:  Grid1.ColWidth(1) = 800: Grid1.Text = "TIPO"
Grid1.Col = 2: Grid1.ColWidth(2) = 1500: Grid1.Text = "VALOR"
Grid1.Col = 3: Grid1.ColWidth(3) = 1500: Grid1.Text = "%PERC."
Grid1.Col = 4: Grid1.ColWidth(4) = 1900: Grid1.Text = "DT.MOVMENTO"
Grid1.Col = 4: Grid1.BackColor = &H80FFFF

Grid1.Row = 0

Grid1.HighLight = False

End Sub

Private Sub txtNome_Change()
Dim nx As Integer
Dim sPesquisa As String
Dim sCritica As String

On Error GoTo Error

If Len(Trim(Label2.Caption)) = 0 Then Exit Sub

sPesquisa = "nome LIKE "
sPesquisa = sPesquisa & "%" & UCase(Trim(Label2.Caption)) & "%"
sCritica = sPesquisa
rs.Filter = sCritica

If rs.RecordCount = 0 Then
   sPesquisa = "NOME > "
   sPesquisa = sPesquisa & Chr$(39) & UCase(Trim(Label2.Caption)) & Chr$(39)
   sPesquisa = sPesquisa & " or NOME="
   sPesquisa = sPesquisa & Chr$(39) & UCase(Trim(Label2.Caption)) & Chr$(39)
   sCritica = sPesquisa
   rs.Filter = sCritica
End If

Call Limpar_Grid
Call Carregar_Grid

Exit Sub

Error:

MsgBox "Nâo digite espacos no campo da pesquisa"

End Sub

Public Function Carregar_Grid()
Dim nx As Double
Dim nLinhas As String
Dim cRec As ADODB.Recordset
Dim sClass As String

Me.Grid1.Visible = False
Me.MousePointer = vbHourglass

Grid1.Row = 1
rs.MoveFirst

'           "CMU_MOV_FUNCIONARIO.MFU_ANO_MES, " & _
'           "CMU_MOV_FUNCIONARIO.MFU_TIPO, " & _
'           "CMU_MOV_FUNCIONARIO.MFU_VALOR, " & _
'           "CMU_MOV_FUNCIONARIO.MFU_PER_DESC, " & _
'           "CMU_MOV_FUNCIONARIO.MFU_DT_MOV "

For nx = 1 To rs.RecordCount
    Grid1.Col = 0: Grid1.Text = Mid$(rs.Fields("MFU_ANO_MES"), 5, 2) & "/" & Mid$(rs.Fields("MFU_ANO_MES"), 1, 4)
    Grid1.Col = 1: Grid1.Text = rs.Fields("MFU_TIPO")
    Grid1.Col = 2: Grid1.Text = Format(rs.Fields("MFU_VALOR"), "#,##0.00")
    Grid1.Col = 3: Grid1.Text = Format(rs.Fields("MFU_PER_DESC"), "#,##0.00")
    Grid1.Col = 4: Grid1.Text = rs.Fields("MFU_DT_MOV")
    rs.MoveNext
    If Not rs.EOF Then
       Grid1.Rows = Grid1.Rows + 1
       Grid1.Row = Grid1.Row + 1
    End If
Next

'Grid1.Col = 1
'Grid1.Sort = flexSortStringAscending

If Grid1.Rows > 2 Then
   Grid1.Rows = Grid1.Rows
   If Len(Trim(Grid1.Text)) = 0 Then Grid1.RemoveItem (Grid1.Row)
   Grid1.Row = 1
   If Len(Trim(Grid1.Text)) = 0 Then Grid1.RemoveItem (Grid1.Row)
   
End If

Me.Grid1.Visible = True
Me.txtlidos.Text = Grid1.Rows - 1
Me.MousePointer = vbDefault
Set cRec = Nothing

Exit Function

Error:

Set cRec = Nothing
Me.MousePointer = vbDefault

End Function




