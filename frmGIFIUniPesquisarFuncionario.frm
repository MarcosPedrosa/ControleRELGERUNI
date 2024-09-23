VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmGIFIUniPesquisarFuncionario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisar funcionário"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   1830
      TabIndex        =   13
      Top             =   3840
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.OptionButton Opt_chapa 
      Caption         =   "Matricula"
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   4770
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.OptionButton Opt_secao 
      Caption         =   "Seção"
      Height          =   255
      Left            =   4650
      TabIndex        =   10
      Top             =   4770
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.OptionButton Opt_nome 
      Caption         =   "Nome"
      Height          =   255
      Left            =   3750
      TabIndex        =   9
      Top             =   4770
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtNome 
      Height          =   315
      Left            =   1830
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   540
      MaxLength       =   6
      TabIndex        =   2
      Top             =   3840
      Width           =   585
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "&Selecionar"
      Height          =   330
      Left            =   7350
      TabIndex        =   1
      Top             =   3840
      Width           =   1275
   End
   Begin VB.CommandButton cmdfechar 
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   8670
      TabIndex        =   0
      Top             =   3840
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   60
      TabIndex        =   6
      Top             =   30
      Width           =   9885
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3375
         Left            =   150
         TabIndex        =   7
         Top             =   240
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   5953
         _Version        =   393216
         Cols            =   4
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
         TabIndex        =   11
         Top             =   1410
         Width           =   5055
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Classificação:"
      Height          =   195
      Left            =   2610
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Nome :"
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   3870
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Lidos: "
      Height          =   285
      Left            =   60
      TabIndex        =   4
      Top             =   3870
      Width           =   465
   End
End
Attribute VB_Name = "frmGIFIUniPesquisarFuncionario"
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

Private Sub cmdSelecionar_Click()
Dim nCod As Double

Me.Grid1.Col = 0
nCod = Me.Grid1.Text
ccodigo_pesquisa = nCod
Me.Grid1.Col = 1
cnome = Me.Grid1.Text
Me.Hide
End Sub

Private Sub Form_Activate()
'   txtNome.SetFocus
'   Me.Grid1.SetFocus
   Me.SetFocus
   nTeclou_Enter = 0
End Sub

Private Sub Form_Load()
'   Me.Show
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

Set rs = CCTempneUniMvFun.MovFuncionario_Consulta_Saldo(cColigada, "")

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
Grid1.Col = 0: Grid1.ColWidth(0) = 900:  Grid1.Text = "CHAPA"
Grid1.Col = 1:  Grid1.ColWidth(1) = 4500: Grid1.Text = "NOME"
Grid1.Col = 2: Grid1.ColWidth(2) = 3500: Grid1.Text = "SEÇÃO"
Grid1.Col = 3: Grid1.ColWidth(3) = 2: Grid1.Text = ""
Grid1.Col = 2: Grid1.BackColor = &H80FFFF

Grid1.Row = 0

Grid1.HighLight = False

End Sub



Private Sub Grid1_DblClick()
Dim nCod As Double

Me.Grid1.Col = 0
nCod = Me.Grid1.Text
ccodigo_pesquisa = nCod
Me.Grid1.Col = 1
cnome = Me.Grid1.Text
Me.Hide
End Sub

Private Sub Opt_chapa_Click()
Call Carrega_januspesquisa
End Sub

Private Sub Opt_Nome_Click()
Call Carrega_januspesquisa
End Sub

Private Sub Opt_secao_Click()
Call Carrega_januspesquisa
End Sub

Private Sub txtNome_Change()
Dim nx As Integer
Dim sPesquisa As String
Dim sCritica As String

On Error GoTo Error

If Len(Trim(txtNome.Text)) = 0 Then Exit Sub

sPesquisa = "NOME LIKE "
sPesquisa = sPesquisa & "%" & UCase(Trim(txtNome.Text)) & "%"
sCritica = sPesquisa
rs.Filter = sCritica

If rs.RecordCount = 0 Then
   sPesquisa = "NOME > "
   sPesquisa = sPesquisa & Chr$(39) & UCase(Trim(txtNome.Text)) & Chr$(39)
   sPesquisa = sPesquisa & " OR NOME = "
   sPesquisa = sPesquisa & Chr$(39) & UCase(Trim(txtNome.Text)) & Chr$(39)
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
Dim cRecAux As ADODB.Recordset
Dim sClass As String
Dim nLidos As Double
Dim sPesquisa As String
Dim sCritica As String

Me.Grid1.Visible = False
Me.MousePointer = vbHourglass

Me.ProgressBar1.Min = 1
Me.ProgressBar1.Max = rs.RecordCount + 2
'Me.ProgressBar1.Value = 0

'Dim cRec As ADODB.Recordset

Set cRec = New ADODB.Recordset
Set cRec = CCTempneTabRegPagto.TabRegPagto_Consultar_Funcionario_Geral(sBancoRM, _
                                                                       cColigada)

Grid1.Row = 1
rs.MoveFirst

For nx = 1 To rs.RecordCount
    Rem ACHAR O FUNCIONARIO NO CADASTRO DE FUNCIONARIOS NA RM
    If rs!SAL_SALDO > 0 Then
       sPesquisa = "CHAPA = "
       sPesquisa = sPesquisa & Chr$(39) & Trim(Format(rs!SAL_CHAPA, "00000")) & Chr$(39)
       sCritica = sPesquisa
       cRec.Filter = sCritica
       If cRec.RecordCount > 0 Then
          nLinhas = Format(cRec!chapa, "000000")
          Grid1.Col = 0: Grid1.Text = nLinhas
          Grid1.Col = 1: Grid1.Text = cRec.Fields("nome")
          Grid1.Col = 2: Grid1.Text = cRec!CODSITUACAO & " - " & cRec.Fields("descricao")
          rs.MoveNext
          If Not rs.EOF Then
             Grid1.Rows = Grid1.Rows + 1
             Grid1.Row = Grid1.Row + 1
          End If
       Else
          rs.MoveNext
       End If
    Else
       rs.MoveNext
    End If
       
    sPesquisa = "CHAPA > "
    sPesquisa = sPesquisa & Chr$(39) & "00000" & Chr$(39)
    sCritica = sPesquisa
    cRec.Filter = sCritica
      
       
       
       
'''       Set cRec = New ADODB.Recordset
'''       Set cRec = CCTempneTabRegPagto.TabRegPagto_Consultar_Funcionario(sBancoRM, _
'''                                                                        cColigada, _
'''                                                                        Format(rs!SAL_CHAPA, "00000"))
'''       If cRec!CODSITUACAO <> "A" And cRec!CODSITUACAO <> "F" Then
'''          nLinhas = Format(cRec.Fields("CHAPA"), "000000")
'''          Grid1.Col = 0: Grid1.Text = nLinhas
'''          Grid1.Col = 1: Grid1.Text = cRec.Fields("nome")
'''          Grid1.Col = 2: Grid1.Text = cRec!CODSITUACAO & " - " & cRec.Fields("descricao")
'''          rs.MoveNext
'''          If Not rs.EOF Then
'''             Grid1.Rows = Grid1.Rows + 1
'''             Grid1.Row = Grid1.Row + 1
'''          End If
'''       Else
'''          rs.MoveNext
'''       End If
'''    Else
'''       rs.MoveNext
'''    End If
    nLidos = nLidos + 1
    ProgressBar1.Value = nLidos
Next

Grid1.Col = 1
Grid1.Sort = flexSortStringAscending

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


