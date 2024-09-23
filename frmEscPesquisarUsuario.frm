VERSION 5.00
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Begin VB.Form frmGIFDiscPesquisarUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisar usuário"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   Icon            =   "frmEscPesquisarUsuario.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNome 
      Height          =   315
      Left            =   870
      TabIndex        =   5
      Top             =   3960
      Width           =   6315
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   1350
      MaxLength       =   6
      TabIndex        =   4
      Top             =   4380
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Height          =   3765
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   7095
      Begin GridEX16.GridEX GEXPesquisa 
         Height          =   3405
         Left            =   150
         TabIndex        =   3
         Top             =   210
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   6006
         MethodHoldFields=   -1  'True
         Options         =   8
         AllowColumnDrag =   0   'False
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         DataMode        =   1
         ColumnHeaderHeight=   285
         IntProp2        =   0
         IntProp7        =   0
      End
   End
   Begin VB.CommandButton cmdSelecionar 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Selecionar"
      Height          =   330
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4350
      Width           =   1275
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4350
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Nome :"
      Height          =   285
      Left            =   150
      TabIndex        =   7
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Total registros : "
      Height          =   285
      Left            =   150
      TabIndex        =   6
      Top             =   4410
      Width           =   1185
   End
End
Attribute VB_Name = "frmGIFDiscPesquisarUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ccodigo_pesquisa As String 'Codigo escolhido pelo usuário.
Public cnome As String 'Nome do escolhido pelo usuário.
Public rs As ADODB.Recordset

Private Sub cmdfechar_Click()
    ccodigo_pesquisa = ""
    cnome = ""
    Unload Me
End Sub

Private Sub cmdSelecionar_Click()
    ccodigo_pesquisa = GEXPesquisa.Value(1)
    cnome = GEXPesquisa.Value(2)
    Me.Hide
End Sub

Private Sub Form_Activate()
   txtNome.SetFocus
End Sub

Private Sub Form_Load()
   Carrega_januspesquisa
End Sub
Private Sub GEXPesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then GEXPesquisa_DblClick
If KeyCode = 27 Then cmdfechar_Click
End Sub
Private Sub GEXPesquisa_Click()
    ccodigo_pesquisa = GEXPesquisa.Value(1)
    cnome = GEXPesquisa.Value(2)
End Sub
Function Carrega_januspesquisa()

Dim nx As Integer

On Error GoTo Erro

Set rs = New ADODB.Recordset

Me.MousePointer = vbHourglass
Set rs = CCTempneUsuario.USUARIO_Consultar(sNomeBanco)
Set GEXPesquisa.ADORecordset = rs
txtlidos.Text = rs.RecordCount
With GEXPesquisa
     .Columns(1).Caption = "Código"
     .Columns(1).Width = TextWidth("wwwwwww")
     .Columns(2).Caption = "nome"
     .Columns(2).Width = TextWidth("wwwwwwwww0wwwwwwwww0wwwwwwwww0")
     For nx = 3 To .Columns.Count
         .Columns(nx).Visible = False
     Next nx
End With
Me.GEXPesquisa.SortKeys.Add 2, jgexSortAscending
Me.MousePointer = vbDefault

Rem Set rs = Nothing
Exit Function

Erro:
MsgBox Err.Description
Me.MousePointer = vbDefault
End Function


Private Sub GEXPesquisa_ColumnHeaderClick(ByVal Column As GridEX16.JSColumn)
If Me.GEXPesquisa.SortKeys.Count > 0 Then
    If Me.GEXPesquisa.SortKeys(1).ColIndex = Column.Index Then
        If Me.GEXPesquisa.SortKeys(1).SortOrder = jgexSortAscending Then
            Me.GEXPesquisa.SortKeys(1).SortOrder = jgexSortDescending
        Else
            Me.GEXPesquisa.SortKeys(1).SortOrder = jgexSortAscending
        End If
        Exit Sub
    End If
    Me.GEXPesquisa.SortKeys.Clear
End If
Me.GEXPesquisa.SortKeys.Add Column.Index, jgexSortAscending

End Sub

Private Sub GEXPesquisa_DblClick()
    ccodigo_pesquisa = GEXPesquisa.Value(1)
    cnome = GEXPesquisa.Value(2)
    Me.Hide
End Sub

Private Sub gexPesquisa_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
Dim nx As Integer
    ccodigo_pesquisa = GEXPesquisa.Value(1)
    cnome = GEXPesquisa.Value(2)
End Sub

Private Sub txtNome_Change()
Dim nx As Integer
Dim sPesquisa As String
Dim sCritica As String

sPesquisa = "USU_USUARIO = "
sPesquisa = sPesquisa & Chr$(39) & UCase(Trim(txtNome.Text)) & Chr$(39)
sCritica = sPesquisa
rs.Filter = sCritica
If rs.RecordCount = 0 Then
   sPesquisa = "USU_USUARIO > "
   sPesquisa = sPesquisa & Chr$(39) & UCase(Trim(txtNome.Text)) & Chr$(39)
   sPesquisa = sPesquisa & " or USU_USUARIO="
   sPesquisa = sPesquisa & Chr$(39) & UCase(Trim(txtNome.Text)) & Chr$(39)
   sCritica = sPesquisa
   rs.Filter = sCritica
Else
   GEXPesquisa_Click
   cmdSelecionar.Default = True
End If
Set GEXPesquisa.ADORecordset = rs
With GEXPesquisa
     .Columns(1).Caption = "Código"
     .Columns(1).Width = TextWidth("wwwwwww")
     .Columns(2).Caption = "Nome"
     .Columns(2).Width = TextWidth("wwwwwwwww0wwwwwwwww0wwwwwwwww0")
     For nx = 3 To .Columns.Count
         .Columns(nx).Visible = False
     Next nx
End With

End Sub






