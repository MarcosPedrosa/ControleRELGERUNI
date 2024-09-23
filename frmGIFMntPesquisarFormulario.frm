VERSION 5.00
Object = "{312C990C-63A1-11D2-ACB5-0080ADA85544}#1.0#0"; "GridEX16.ocx"
Begin VB.Form frmGIFMntPesquisarFormulario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisar Formul�rios do Sistema"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNome 
      Height          =   315
      Left            =   840
      TabIndex        =   5
      Top             =   3930
      Width           =   7935
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   4
      Top             =   4350
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Height          =   3765
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   8925
      Begin GridEX16.GridEX GEXPesquisa 
         Height          =   3405
         Left            =   150
         TabIndex        =   3
         Top             =   210
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   6006
         HeaderStyle     =   2
         UseEvenOddColor =   -1  'True
         MethodHoldFields=   -1  'True
         Options         =   8
         AllowColumnDrag =   0   'False
         RecordsetType   =   1
         AllowEdit       =   0   'False
         GroupByBoxVisible=   0   'False
         RowHeaders      =   -1  'True
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
      Left            =   6420
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4350
      Width           =   1275
   End
   Begin VB.CommandButton cmdfechar 
      BackColor       =   &H000000FF&
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   7740
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4350
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Nome :"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3990
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Total registros : "
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   4380
      Width           =   1185
   End
End
Attribute VB_Name = "frmGIFMntPesquisarFormulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ccodigo_pesquisa As String 'Codigo escolhido pelo usu�rio.
Public cnome As String 'Nome do escolhido pelo usu�rio.
Public cTipo As Integer 'Tipo da tabela de acordo com as opcoes abaixo
'1 - tabela de marca de veiculos
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
   Me.txtNome.SetFocus
End Sub

Private Sub Form_Load()
   Call Carrega_januspesquisa
End Sub
Private Sub GEXPesquisa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Me.GEXPesquisa.MovePrevious
   GEXPesquisa_DblClick
End If
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

Set rs = CCTempneFormulario.TipoForm_Consultar(sNomeBanco)

Set GEXPesquisa.ADORecordset = rs
txtlidos.Text = rs.RecordCount

With GEXPesquisa
     .Columns(1).Caption = "C�digo"
     .Columns(1).Width = TextWidth("wwwww")
     .Columns(2).Caption = "Menu"
     .Columns(2).Width = TextWidth("wwwwwwwwwww0")
     .Columns(3).Caption = "Formul�rio"
     .Columns(3).Width = TextWidth("wwwwWWWWWWWWwwwwww0")
     .Columns(4).Caption = "A��o"
     .Columns(4).Width = TextWidth("wwwwwwWWwwwwww0")
     .Columns(5).Caption = "Grupo"
     .Columns(5).Width = TextWidth("wwwwwwww0")
     For nx = 6 To .Columns.Count
         .Columns(nx).Visible = False
     Next nx
End With
'Me.GEXPesquisa.SortKeys.Add 2, jgexSortAscending
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

sPesquisa = "FOR_NOME_EDITOR = "
sPesquisa = sPesquisa & Chr$(39) & UCase(Trim(txtNome.Text)) & Chr$(39)
sCritica = sPesquisa
rs.Filter = sCritica
If rs.RecordCount = 0 Then
   sPesquisa = "FOR_NOME_EDITOR > "
   sPesquisa = sPesquisa & Chr$(39) & UCase(Trim(txtNome.Text)) & Chr$(39)
   sPesquisa = sPesquisa & " or FOR_NOME_EDITOR="
   sPesquisa = sPesquisa & Chr$(39) & UCase(Trim(txtNome.Text)) & Chr$(39)
   sCritica = sPesquisa
   rs.Filter = sCritica
Else
   GEXPesquisa_Click
   cmdSelecionar.Default = True
End If
Set GEXPesquisa.ADORecordset = rs
With GEXPesquisa
     .Columns(1).Caption = "C�digo"
     .Columns(1).Width = TextWidth("wwwww")
     .Columns(2).Caption = "Menu"
     .Columns(2).Width = TextWidth("wwwwwwwwwww0")
     .Columns(3).Caption = "Formul�rio"
     .Columns(3).Width = TextWidth("wwwwWWWWWWWWwwwwww0")
     .Columns(4).Caption = "A��o"
     .Columns(4).Width = TextWidth("wwwwwwWWwwwwww0")
     .Columns(5).Caption = "Grupo"
     .Columns(5).Width = TextWidth("wwwwwwww0")
     For nx = 6 To .Columns.Count
         .Columns(nx).Visible = False
     Next nx
End With

End Sub

Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Then
   Me.GEXPesquisa.SetFocus
End If

End Sub

